import importlib.util

if importlib.util.find_spec("botbuilder") is None:
    raise ValueError(
        "The botbuilder-core package is required to integrate Chainlit with a Teams app. Run `pip install botbuilder-core --upgrade`"
    )
