import asyncio
import base64
import mimetypes
import os
import uuid
from datetime import datetime
from typing import TYPE_CHECKING, Dict, List, Literal, Optional, Union

import filetype

if TYPE_CHECKING:
    from microsoft_agents.hosting.core import TurnContext
    from microsoft_agents.activity import Activity

import httpx
from microsoft_agents.activity import (
    ActionTypes,
    Activity,
    ActivityTypes,
    Attachment,
    CardAction,
    ChannelAccount,
    HeroCard,
)
from microsoft_agents.hosting.core import (
    MessageFactory,
    RestChannelServiceClientFactory,
    TurnContext,
)
from microsoft_agents.hosting.core.authorization import (
    AgentAuthConfiguration,
    AuthTypes,
    ClaimsIdentity,
    JwtTokenValidator,
)
from microsoft_agents.hosting.core.http import HttpAdapterBase, HttpResponseFactory

from chainlit.config import config
from chainlit.context import ChainlitContext, HTTPSession, context, context_var
from chainlit.data import get_data_layer
from chainlit.element import Element, ElementDict
from chainlit.emitter import BaseChainlitEmitter
from chainlit.logger import logger
from chainlit.message import Message, StepDict
from chainlit.types import Feedback
from chainlit.user import PersistedUser, User
from chainlit.user_session import user_session


class MsAgentsEmitter(BaseChainlitEmitter):
    def __init__(self, session: HTTPSession, turn_context: TurnContext):
        super().__init__(session)
        self.turn_context = turn_context

    async def send_element(self, element_dict: ElementDict):
        if element_dict.get("display") != "inline":
            return

        persisted_file = self.session.files.get(element_dict.get("chainlitKey") or "")
        attachment: Optional[Attachment] = None
        mime: Optional[str] = None

        element_name: str = element_dict.get("name", "Untitled")

        if mime:
            file_extension = mimetypes.guess_extension(mime)
            if file_extension:
                element_name += file_extension

        if persisted_file:
            mime = element_dict.get("mime")
            with open(persisted_file["path"], "rb") as file:
                dencoded_string = base64.b64encode(file.read()).decode()
                content_url = f"data:{mime};base64,{dencoded_string}"
                attachment = Attachment(
                    content_type=mime, content_url=content_url, name=element_name
                )

        elif url := element_dict.get("url"):
            attachment = Attachment(
                content_type=mime, content_url=url, name=element_name
            )

        if not attachment:
            return

        await self.turn_context.send_activity(Activity(type=ActivityTypes.message, attachments=[attachment]))

    async def send_step(self, step_dict: StepDict):
        if not step_dict["type"] == "assistant_message":
            return

        step_type = step_dict.get("type")
        is_message = step_type in [
            "user_message",
            "assistant_message",
        ]
        is_empty_output = not step_dict.get("output")

        if is_empty_output or not is_message:
            return
        else:
            reply = MessageFactory.text(step_dict["output"])
            enable_feedback = get_data_layer()
            if enable_feedback:
                current_run = context.current_run
                scorable_id = current_run.id if current_run else step_dict["id"]
                like_button = CardAction(
                    type=ActionTypes.message_back,
                    title="\U0001f44d",
                    text="like",
                    value={"feedback": "like", "step_id": scorable_id},
                )
                dislike_button = CardAction(
                    type=ActionTypes.message_back,
                    title="\U0001f44e",
                    text="dislike",
                    value={"feedback": "dislike", "step_id": scorable_id},
                )
                card = HeroCard(buttons=[like_button, dislike_button])
                attachment = Attachment(
                    content_type="application/vnd.microsoft.card.hero",
                    content=card.model_dump(by_alias=True, exclude_none=True),
                )
                reply.attachments = [attachment]

            await self.turn_context.send_activity(reply)

    async def update_step(self, step_dict: StepDict):
        if not step_dict["type"] == "assistant_message":
            return

        await self.send_step(step_dict)


class _BotTokenProvider:
    """MSAL-based token provider for outbound Bot Framework calls."""

    def __init__(self, config: AgentAuthConfiguration):
        import msal

        self._app = msal.ConfidentialClientApplication(
            client_id=config.CLIENT_ID,
            client_credential=config.CLIENT_SECRET,
            authority=config.AUTHORITY
            or f"https://login.microsoftonline.com/{config.TENANT_ID}",
        )

    async def get_access_token(
        self, resource_url: str, scopes: list, force_refresh: bool = False
    ) -> str:
        scope = scopes if scopes else [f"{resource_url}/.default"]
        result = None
        if not force_refresh:
            result = self._app.acquire_token_silent(scope, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=scope)
        if "access_token" in result:
            return result["access_token"]
        raise ValueError(
            f"Failed to acquire token: {result.get('error_description', result.get('error', 'unknown'))}"
        )

    async def acquire_token_on_behalf_of(self, scopes: list, user_assertion: str) -> str:
        raise NotImplementedError()

    async def get_agentic_application_token(self, agent_app_instance_id: str):
        raise NotImplementedError()

    async def get_agentic_instance_token(self, agent_app_instance_id: str):
        raise NotImplementedError()

    async def get_agentic_user_token(
        self, agent_app_instance_id: str, agentic_user_id: str, scopes: list
    ):
        raise NotImplementedError()


class _BotConnections:
    """Simple Connections implementation for a single-tenant bot."""

    def __init__(self, config: AgentAuthConfiguration):
        self._config = config
        self._provider = _BotTokenProvider(config)

    def get_connection(self, connection_name: str):
        return self._provider

    def get_default_connection(self):
        return self._provider

    def get_token_provider(self, claims_identity, service_url: str):
        return self._provider

    def get_default_connection_configuration(self):
        return self._config


class _StarletteRequestAdapter:
    """Adapts a Starlette/FastAPI Request to HttpRequestProtocol for use with HttpAdapterBase."""

    def __init__(self, request, claims_identity=None):
        self._request = request
        self._claims = claims_identity

    @property
    def method(self) -> str:
        return self._request.method

    @property
    def headers(self):
        return self._request.headers

    async def json(self):
        return await self._request.json()

    def get_claims_identity(self):
        return self._claims

    def get_path_param(self, name: str) -> str:
        return self._request.path_params.get(name, "")


class CloudAdapter(HttpAdapterBase):
    """FastAPI/Starlette-compatible CloudAdapter for the M365 Agents SDK."""

    def __init__(self, auth_config: AgentAuthConfiguration):
        self._auth_config = auth_config
        self._token_validator = JwtTokenValidator(auth_config)
        connections = _BotConnections(auth_config)
        factory = RestChannelServiceClientFactory(connections)
        super().__init__(channel_service_client_factory=factory)

    async def process(self, request, agent):
        """Process a FastAPI/Starlette request and return an HttpResponse."""
        auth_header = request.headers.get("Authorization", "")
        claims_identity = None

        if auth_header and auth_header.startswith("Bearer "):
            token = auth_header.split(" ")[1]
            try:
                claims_identity = await self._token_validator.validate_token(token)
            except ValueError:
                return HttpResponseFactory.unauthorized()
        elif self._auth_config.CLIENT_ID:
            return HttpResponseFactory.unauthorized()

        adapted_req = _StarletteRequestAdapter(request, claims_identity)
        return await self.process_request(adapted_req, agent)


_auth_config = AgentAuthConfiguration(
    auth_type=AuthTypes.client_secret,
    client_id=os.environ.get("MICROSOFT_APP_ID"),
    tenant_id=os.environ.get("MICROSOFT_APP_TENANT_ID"),
    client_secret=os.environ.get("MICROSOFT_APP_PASSWORD"),
)
adapter = CloudAdapter(_auth_config)


def init_msagents_context(
    session: HTTPSession,
    turn_context: TurnContext,
) -> ChainlitContext:
    emitter = MsAgentsEmitter(session=session, turn_context=turn_context)
    context = ChainlitContext(session=session, emitter=emitter)
    context_var.set(context)
    user_session.set("msagents_turn_context", turn_context)
    return context


users_by_msagents_id: Dict[str, Union[User, PersistedUser]] = {}

USER_PREFIX = "msagents_"


async def get_user(msagents_user: ChannelAccount):
    if msagents_user.id in users_by_msagents_id:
        return users_by_msagents_id[msagents_user.id]

    metadata = {
        "name": msagents_user.name,
        "id": msagents_user.id,
    }
    user = User(identifier=USER_PREFIX + str(msagents_user.name), metadata=metadata)

    users_by_msagents_id[msagents_user.id] = user

    if data_layer := get_data_layer():
        try:
            persisted_user = await data_layer.create_user(user)
            if persisted_user:
                users_by_msagents_id[msagents_user.id] = persisted_user
        except Exception as e:
            logger.error(f"Error creating user: {e}")

    return users_by_msagents_id[msagents_user.id]


async def download_msagents_file(url: str):
    async with httpx.AsyncClient() as client:
        response = await client.get(url)
        if response.status_code == 200:
            return response.content
        else:
            return None


async def download_msagents_files(
    session: HTTPSession, attachments: Optional[List[Attachment]] = None
):
    if not attachments:
        return []

    attachments = [
        attachment for attachment in attachments if isinstance(attachment.content, dict)
    ]
    download_coros = [
        download_msagents_file(attachment.content.get("downloadUrl"))
        for attachment in attachments
    ]
    file_bytes_list = await asyncio.gather(*download_coros)
    file_refs = []
    for idx, file_bytes in enumerate(file_bytes_list):
        if file_bytes:
            name = attachments[idx].name
            mime_type = filetype.guess_mime(file_bytes) or "application/octet-stream"
            file_ref = await session.persist_file(
                name=name, mime=mime_type, content=file_bytes
            )
            file_refs.append(file_ref)

    files_dicts = [
        session.files[file["id"]] for file in file_refs if file["id"] in session.files
    ]

    elements = [
        Element.from_dict(
            {
                "id": file["id"],
                "name": file["name"],
                "path": str(file["path"]),
                "chainlitKey": file["id"],
                "display": "inline",
                "type": Element.infer_type_from_mime(file["type"]),
            }
        )
        for file in files_dicts
    ]

    return elements


def clean_content(activity: Activity):
    return activity.text.strip()


async def process_msagents_message(
    turn_context: TurnContext,
    thread_name: str,
):
    user = await get_user(turn_context.activity.from_property)

    thread_id = str(
        uuid.uuid5(
            uuid.NAMESPACE_DNS,
            str(
                turn_context.activity.conversation.id
                + datetime.today().strftime("%Y-%m-%d")
            ),
        )
    )

    text = clean_content(turn_context.activity)
    files = turn_context.activity.attachments

    session_id = str(uuid.uuid4())

    session = HTTPSession(
        id=session_id,
        thread_id=thread_id,
        user=user,
        client_type="msagents",
    )

    ctx = init_msagents_context(
        session=session,
        turn_context=turn_context,
    )

    file_elements = await download_msagents_files(session, files)

    if on_chat_start := config.code.on_chat_start:
        await on_chat_start()

    msg = Message(
        content=text,
        elements=file_elements,
        type="user_message",
        author=user.metadata.get("name"),
    )

    await msg.send()

    if on_message := config.code.on_message:
        await on_message(msg)

    if on_chat_end := config.code.on_chat_end:
        await on_chat_end()

    if data_layer := get_data_layer():
        if isinstance(user, PersistedUser):
            try:
                await data_layer.update_thread(
                    thread_id=thread_id,
                    name=thread_name,
                    metadata=ctx.session.to_persistable(),
                    user_id=user.id,
                )
            except Exception as e:
                logger.error(f"Error updating thread: {e}")

    await ctx.session.delete()


async def handle_message(turn_context: TurnContext):
    if turn_context.activity.type == ActivityTypes.message:
        if (
            turn_context.activity.text == "like"
            or turn_context.activity.text == "dislike"
        ):
            feedback_value: Literal[0, 1] = (
                0 if turn_context.activity.text == "dislike" else 1
            )
            step_id = turn_context.activity.value.get("step_id")
            if data_layer := get_data_layer():
                await data_layer.upsert_feedback(
                    Feedback(forId=step_id, value=feedback_value)
                )
            updated_text = "\U0001f44d" if turn_context.activity.text == "like" else "\U0001f44e"
            # Update the existing message to remove the buttons
            updated_message = Activity(
                type=ActivityTypes.message,
                id=turn_context.activity.reply_to_id,
                text=updated_text,
                attachments=[],
            )
            await turn_context.update_activity(updated_message)
        else:
            # Send typing activity
            typing_activity = Activity(
                type=ActivityTypes.typing,
                from_property=turn_context.activity.recipient,
                recipient=turn_context.activity.from_property,
                conversation=turn_context.activity.conversation,
            )
            await turn_context.send_activity(typing_activity)
            thread_name = f"{turn_context.activity.from_property.name} Teams DM {datetime.today().strftime('%Y-%m-%d')}"
            await process_msagents_message(turn_context, thread_name)


async def on_turn(turn_context: TurnContext):
    await handle_message(turn_context)


class MsAgentsBot:
    async def on_turn(self, turn_context: TurnContext):
        await on_turn(turn_context)


bot = MsAgentsBot()
