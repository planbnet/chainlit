# Microsoft Teams Integration

This guide explains how to connect an existing Chainlit application to Microsoft Teams as a bot. Your `@cl.on_message`, `@cl.on_chat_start`, and `@cl.on_chat_end` callbacks work in Teams with zero code changes.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Architecture Overview](#architecture-overview)
- [Azure Setup](#azure-setup)
- [Chainlit Configuration](#chainlit-configuration)
- [Running Your Bot](#running-your-bot)
- [Publishing to Teams](#publishing-to-teams)
- [Features](#features)
- [Platform-Specific Logic](#platform-specific-logic)
- [Data Persistence & Feedback](#data-persistence--feedback)
- [Limitations](#limitations)
- [Troubleshooting](#troubleshooting)
- [TODO / Roadmap](#todo--roadmap)

---

## Prerequisites

### Python Packages

Install the Microsoft 365 Agents SDK packages and MSAL alongside Chainlit:

```bash
pip install chainlit microsoft-agents-hosting-core microsoft-agents-hosting-aiohttp msal
```

| Package | Version | Purpose |
|---|---|---|
| `microsoft-agents-hosting-core` | >= 0.7.0 | Core bot framework (TurnContext, Activity model, JWT validation) |
| `microsoft-agents-hosting-aiohttp` | >= 0.7.0 | HTTP adapter layer used internally |
| `msal` | >= 1.20.0 | Azure AD token acquisition (client credentials flow) |

> **Note:** The older `botbuilder-python` SDK reached end-of-life on December 31, 2025 and has been archived. Chainlit uses its successor, the [Microsoft 365 Agents SDK for Python](https://github.com/microsoft/Agents-for-python).

### Azure Resources

You need the following Azure resources:

1. **Microsoft Entra ID (Azure AD) App Registration** - provides the app identity
2. **Azure Bot resource** - connects your app to the Teams channel
3. A **publicly reachable HTTPS endpoint** for your Chainlit server (Azure App Service, a VM with a reverse proxy, or a tunnel like [ngrok](https://ngrok.com/) / [VS Code dev tunnels](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/) for local development)

---

## Architecture Overview

When a user sends a message in Teams, the flow is:

```
Teams Client
    |
    v
Azure Bot Service  (routes messages to your endpoint)
    |
    v
POST /api/messages  (Chainlit's FastAPI server)
    |
    v
CloudAdapter  (validates JWT, deserializes Activity)
    |
    v
TeamsBot.on_turn()
    |
    v
process_teams_message()
    |-- Creates an HTTPSession (not a WebSocket session)
    |-- Calls @cl.on_chat_start
    |-- Creates a cl.Message from the user's text + attachments
    |-- Calls @cl.on_message
    |-- Calls @cl.on_chat_end
    |-- Persists the thread (if a data layer is configured)
    v
TeamsEmitter  (sends replies back via TurnContext)
```

Your existing Chainlit callbacks are called exactly as they would be from the web UI. The `TeamsEmitter` translates `cl.Message` outputs into Teams Activity messages automatically.

---

## Azure Setup

### Step 1: Register an App in Microsoft Entra ID

1. Go to the [Azure Portal](https://portal.azure.com/) > **Microsoft Entra ID** > **App registrations** > **New registration**.
2. Set a name (e.g., "Chainlit Bot").
3. For **Supported account types**, choose the appropriate option:
   - *Single tenant* for internal/org-only bots
   - *Multitenant* if you plan to distribute the bot
4. Leave **Redirect URI** blank. Click **Register**.
5. On the app's **Overview** page, copy:
   - **Application (client) ID** &rarr; this is your `MICROSOFT_APP_ID`
   - **Directory (tenant) ID** &rarr; this is your `MICROSOFT_APP_TENANT_ID`
6. Go to **Certificates & secrets** > **New client secret**. Copy the secret value &rarr; this is your `MICROSOFT_APP_PASSWORD`.

### Step 2: Create an Azure Bot Resource

1. In the Azure Portal, search for **Azure Bot** and create a new resource.
2. Under **Bot handle**, enter a unique name.
3. Under **Type of App**, select the matching option for your registration (Single Tenant or Multi Tenant).
4. Paste your **Application (client) ID** from Step 1.
5. Once created, go to the bot resource > **Channels** > **Microsoft Teams** and enable it.
6. Go to **Configuration** and set the **Messaging endpoint** to:

   ```
   https://<your-public-domain>/api/messages
   ```

   For local development with ngrok:

   ```
   https://<random-id>.ngrok-free.app/api/messages
   ```

### Step 3: Set Environment Variables

Set these three environment variables in your deployment or `.env` file:

```bash
MICROSOFT_APP_ID=<your-application-client-id>
MICROSOFT_APP_PASSWORD=<your-client-secret-value>
MICROSOFT_APP_TENANT_ID=<your-directory-tenant-id>
```

Chainlit checks for `MICROSOFT_APP_ID` and `MICROSOFT_APP_PASSWORD` at startup. When both are present, the Teams integration is automatically activated and the `POST /api/messages` endpoint is registered.

---

## Chainlit Configuration

### No Code Changes Required

Your existing Chainlit app works as-is. For example:

```python
import chainlit as cl

@cl.on_chat_start
async def start():
    await cl.Message(content="Hello from Teams!").send()

@cl.on_message
async def main(message: cl.Message):
    await cl.Message(content=f"You said: {message.content}").send()

@cl.on_chat_end
async def end():
    # Cleanup resources if needed
    pass
```

All three callbacks (`on_chat_start`, `on_message`, `on_chat_end`) are invoked for every incoming Teams message. The lifecycle is per-message (not a persistent WebSocket connection), so `on_chat_start` runs at the beginning of each request and `on_chat_end` at the end.

### Headless Mode (Optional)

If your deployment is Teams-only and you don't want to serve the web UI, use headless mode:

```bash
chainlit run app.py --headless
```

This prevents the browser from opening but still serves the FastAPI application with the `/api/messages` endpoint.

---

## Running Your Bot

### Local Development

1. Start your Chainlit app:

   ```bash
   chainlit run app.py
   ```

   By default this listens on `http://127.0.0.1:8000`.

2. Expose the local server with ngrok (or another tunnel):

   ```bash
   ngrok http 8000
   ```

3. Copy the ngrok HTTPS URL and update your Azure Bot's messaging endpoint:

   ```
   https://<id>.ngrok-free.app/api/messages
   ```

4. Open Teams, find your bot (by the name from your Azure Bot resource), and send a message.

### Production Deployment

Deploy your Chainlit app to any hosting platform that provides a stable HTTPS URL (Azure App Service, AWS, GCP, a VM with Nginx/Caddy, etc.) and set the messaging endpoint in your Azure Bot configuration to match.

Example with a custom host and port:

```bash
chainlit run app.py --host 0.0.0.0 --port 8000
```

---

## Publishing to Teams

### Teams App Manifest

To make your bot discoverable in Teams, you need a Teams app manifest. Chainlit includes a template at `backend/chainlit/teams/manifest_template.json`:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
  "manifestVersion": "1.17",
  "version": "1.0.0",
  "id": "<YOUR_ENTRA_APP_CLIENT_ID>",
  "packageName": "com.example.chainlit-bot",
  "developer": {
    "name": "Your Organization",
    "websiteUrl": "https://your-app.azurewebsites.net",
    "privacyUrl": "https://your-app.azurewebsites.net/privacy",
    "termsOfUseUrl": "https://your-app.azurewebsites.net/terms"
  },
  "icons": {
    "color": "color.png",
    "outline": "outline.png"
  },
  "name": {
    "short": "Chainlit Bot",
    "full": "Chainlit Bot"
  },
  "description": {
    "short": "Chainlit AI assistant",
    "full": "Chainlit-powered AI assistant integrated with Microsoft Teams"
  },
  "accentColor": "#FFFFFF",
  "bots": [
    {
      "botId": "<YOUR_ENTRA_APP_CLIENT_ID>",
      "scopes": ["personal", "team", "groupChat"],
      "supportsFiles": false,
      "isNotificationOnly": false
    }
  ],
  "permissions": ["identity", "messageTeamMembers"],
  "validDomains": ["your-app.azurewebsites.net"]
}
```

To create the installable package:

1. Copy the template and replace `<YOUR_ENTRA_APP_CLIENT_ID>` with your app's client ID.
2. Update `developer`, `name`, `description`, and `validDomains` fields.
3. Add two icon files: `color.png` (192x192) and `outline.png` (32x32, transparent).
4. Zip the manifest JSON and both icons into a `.zip` file.
5. In Teams Admin Center or Teams Developer Portal, upload the zip as a custom app.

### Bot Scopes

The manifest supports three scopes:

| Scope | Description |
|---|---|
| `personal` | 1:1 direct messages with the bot |
| `team` | Bot added to a team channel |
| `groupChat` | Bot added to a group chat |

---

## Features

### What Works

| Feature | Status | Details |
|---|---|---|
| Text messages | Supported | Full send/receive via `@cl.on_message` |
| File uploads (user &rarr; bot) | Supported | Files are downloaded, persisted, and passed as `cl.Element` objects in `message.elements` |
| File attachments (bot &rarr; user) | Supported | Inline elements are sent as base64-encoded attachments or URL attachments |
| Typing indicators | Supported | A typing activity is sent when the bot begins processing a message |
| Feedback (thumbs up/down) | Supported | HeroCard buttons appended to bot replies when a data layer is configured |
| Thread management | Supported | Conversations are grouped into daily threads per conversation ID |
| User identification | Supported | Teams user name and ID are extracted and persisted |
| `@cl.on_chat_start` | Supported | Called at the start of each message processing |
| `@cl.on_chat_end` | Supported | Called after message processing completes |
| `cl.user_session` | Supported | Available during the request lifecycle |

### Typing Indicator

When a user sends a message, the bot automatically sends a typing indicator before processing begins. This gives the user visual feedback that the bot is working on a response.

### File Handling

**Receiving files from users:** When a user attaches files to a Teams message, Chainlit downloads them, detects their MIME type, and makes them available as `cl.Element` objects on the incoming `cl.Message`:

```python
@cl.on_message
async def main(message: cl.Message):
    if message.elements:
        for element in message.elements:
            # element.name, element.path, element.mime are available
            await cl.Message(content=f"Received file: {element.name}").send()
    else:
        await cl.Message(content=f"Echo: {message.content}").send()
```

**Sending files to users:** Elements with `display="inline"` are sent as Teams attachments (base64-encoded or as a URL).

### Feedback Buttons

When a [data layer](https://docs.chainlit.io/data-persistence/overview) is configured, every assistant message includes a HeroCard with thumbs-up and thumbs-down buttons. When a user clicks one:

1. The feedback is persisted via `data_layer.upsert_feedback()`
2. The original message is updated to show the selected emoji and the buttons are removed

---

## Platform-Specific Logic

You can detect whether a request is coming from Teams using `cl.user_session`:

```python
import chainlit as cl

@cl.on_message
async def main(message: cl.Message):
    client_type = cl.user_session.get("client_type")

    if client_type == "teams":
        # Teams-specific behavior
        await cl.Message(content="Hello from Teams!").send()
    else:
        # Web UI or other platform
        await cl.Message(content="Hello from the web!").send()
```

The `client_type` value is `"teams"` for Teams messages. Other possible values are `"webapp"`, `"copilot"`, `"slack"`, and `"discord"`.

### Accessing the TurnContext

For advanced Teams-specific operations, you can access the raw `TurnContext` from the Microsoft 365 Agents SDK:

```python
@cl.on_message
async def main(message: cl.Message):
    turn_context = cl.user_session.get("teams_turn_context")
    if turn_context:
        # Access raw Teams activity data
        conversation_id = turn_context.activity.conversation.id
        user_name = turn_context.activity.from_property.name
```

---

## Data Persistence & Feedback

### Configuring a Data Layer

To enable thread persistence and feedback collection, configure a [Chainlit data layer](https://docs.chainlit.io/data-persistence/overview). When a data layer is present:

- **Threads** are created/updated in the data layer after each message, named with the pattern `"{UserName} Teams DM {YYYY-MM-DD}"`.
- **Thread IDs** are deterministic per conversation per day (using UUID5 from the Teams conversation ID + current date), so messages within the same Teams conversation on the same day belong to the same thread.
- **Users** are created in the data layer with identifiers prefixed with `teams_` (e.g., `teams_John Doe`).
- **Feedback** from HeroCard buttons is stored via `data_layer.upsert_feedback()`.

### Without a Data Layer

Without a data layer, the bot still functions normally for sending and receiving messages. Feedback buttons will not appear, and conversation history will not be persisted.

---

## Limitations

| Limitation | Details |
|---|---|
| No streaming | Teams does not support streaming token-by-token responses. The full response is sent as a single message. |
| No Adaptive Cards (yet) | Only HeroCards are used for feedback. Custom Adaptive Card support is not built in. |
| No proactive messaging | The bot can only respond to incoming messages. Proactive outbound messaging is not supported. |
| No `@cl.on_chat_resume` | Since Teams uses HTTP sessions (not WebSocket), chat resume is not available. `on_chat_start` is called on every message. |
| No chat profiles | Chat profiles and starters are web UI features and do not apply to Teams. |
| No audio/video | Teams audio and video features are not bridged to Chainlit's audio hooks. |
| No message editing by user | Teams message edits are not forwarded as new events. |
| No `@cl.action_callback` | Action buttons from `cl.Action` are web-UI-specific. Use the built-in feedback HeroCard buttons instead. |
| Session is per-message | Unlike the web UI where a session persists across a conversation, each Teams message creates a new `HTTPSession`. Data stored in `cl.user_session` does not carry over between messages. |

---

## Troubleshooting

### Bot not responding

1. **Check environment variables** - Ensure `MICROSOFT_APP_ID`, `MICROSOFT_APP_PASSWORD`, and `MICROSOFT_APP_TENANT_ID` are set correctly.
2. **Check the messaging endpoint** - In the Azure Bot resource > Configuration, verify the endpoint URL points to `https://<your-domain>/api/messages`.
3. **Check the server logs** - Look for errors in the Chainlit console output. Auth failures will show as 401 responses.
4. **Check the tunnel** - If using ngrok, verify it's running and the URL matches the bot configuration.

### 401 Unauthorized

- The JWT token validation failed. This usually means:
  - `MICROSOFT_APP_ID` or `MICROSOFT_APP_PASSWORD` is incorrect
  - `MICROSOFT_APP_TENANT_ID` doesn't match the app registration
  - The client secret has expired (they expire after 6, 12, or 24 months)

### Bot responds in development but not in Teams

- Ensure the Teams channel is enabled on the Azure Bot resource.
- Ensure the app is installed/sideloaded in your Teams tenant.
- Check that your domain is listed in `validDomains` in the manifest.

### Import errors on startup

If you see `The microsoft-agents-hosting-core package is required to integrate Chainlit with a Teams app`, install the required packages:

```bash
pip install microsoft-agents-hosting-core microsoft-agents-hosting-aiohttp msal
```

### Missing feedback buttons

Feedback buttons only appear when a data layer is configured. Without one, messages are sent as plain text.

---

## TODO / Roadmap

The following features are planned or could be contributed:

- [ ] **Adaptive Card support** - Render rich Adaptive Cards instead of plain text for structured responses (tables, forms, etc.)
- [ ] **Proactive messaging** - Allow the bot to send messages to users without a prior incoming message (notifications, alerts)
- [ ] **Message streaming simulation** - Update a single message multiple times to simulate streaming (using `update_activity`)
- [ ] **Conversation state persistence** - Carry `cl.user_session` data across messages within the same Teams conversation using the data layer
- [ ] **`@mention` support in channels** - Respond only when `@mentioned` in team channels, with proper mention entity parsing
- [ ] **Adaptive Card actions** - Map `cl.Action` callbacks to Adaptive Card Action.Submit buttons
- [ ] **Rich text formatting** - Convert Markdown in bot responses to Teams-compatible formatting
- [ ] **Multi-tenant bot support** - Guide and test multi-tenant app registration for distributing the bot across organizations
- [ ] **File upload (`supportsFiles: true`)** - Enable the Teams file consent flow for the bot to request file upload permission
- [ ] **Bot authentication with SSO** - Support Teams Single Sign-On to map Teams identity to Chainlit auth
- [ ] **Rate limiting / throttling** - Handle Teams throttling (HTTP 429) with retry logic
- [ ] **Unit & integration tests** - Expand test coverage for the Teams integration module
