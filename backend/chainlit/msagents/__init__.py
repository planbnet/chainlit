import importlib.util

if importlib.util.find_spec("microsoft_agents") is None:
    raise ValueError(
        "The microsoft-agents-hosting-core package is required to integrate Chainlit with Microsoft 365 Agents. "
        "Run `pip install microsoft-agents-hosting-core microsoft-agents-hosting-aiohttp msal`"
    )
