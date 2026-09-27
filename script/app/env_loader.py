"""
Utility module for loading environment variables from .env file.
"""

import os
from pathlib import Path
from typing import Any

import dotenv
from openai import AzureOpenAI


# Requested startup style: load env at import/startup.
dotenv.load_dotenv()

endpoint = ""
model_name = "gpt-4.1-mini"
deployment = ""
api_version = "2024-12-01-preview"


def _clean(value: str | None, default: str = "") -> str:
    """Normalize env values by trimming spaces and optional surrounding quotes."""
    raw = (value if value is not None else default).strip()
    if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in {'"', "'"}:
        return raw[1:-1].strip()
    return raw


def load_env() -> None:
    """Load environment variables from .env file if it exists.

    Uses python-dotenv to load variables from .env into os.environ.
    If .env does not exist, silently continues (user can enter values manually).
    """
    try:
        from dotenv import load_dotenv
    except ImportError:
        # python-dotenv not installed; skip loading from .env
        return

    env_file = Path(__file__).resolve().parent.parent / ".env"
    if env_file.exists():
        load_dotenv(env_file, override=True)


def get_azure_client(**overrides: Any) -> AzureOpenAI:
    """Create an AzureOpenAI client from settings and optional overrides.

    Parameters
    ----------
    **overrides
        Optional keyword overrides for endpoint, api_key, and api_version.
    """
    settings = get_azure_settings()
    return AzureOpenAI(
        api_version=overrides.get("api_version", settings["api_version"]),
        azure_endpoint=overrides.get("endpoint", settings["endpoint"]),
        api_key=overrides.get("api_key", settings["api_key"]),
    )


def get_azure_settings() -> dict:
    """Retrieve Azure OpenAI settings from environment variables.

    Returns
    -------
    dict
        Dictionary with keys: endpoint, api_key, chat_deployment,
        embed_deployment, api_version (values may be empty strings if not set).
    """
    load_env()
    return {
        "endpoint": _clean(os.getenv("AZURE_ENDPOINT", os.getenv("AZURE_OPENAI_ENDPOINT", endpoint)), endpoint),
        "api_key": _clean(os.getenv("AZURE_API_KEY", os.getenv("AZURE_OPENAI_KEY", "")), ""),
        "chat_deployment": _clean(
            os.getenv("AZURE_CHAT_DEPLOYMENT", os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", deployment)),
            deployment,
        ),
        "embed_deployment": _clean(os.getenv("AZURE_EMBED_DEPLOYMENT", "text-embedding-3-small"), "text-embedding-3-small"),
        "api_version": _clean(
            os.getenv("AZURE_API_VERSION", os.getenv("AZURE_OPENAI_API_VERSION", api_version)),
            api_version,
        ),
        "model_name": _clean(os.getenv("AZURE_OPENAI_MODEL_NAME", model_name), model_name),
        "fernet_key": _clean(os.getenv("FERNET_KEY", ""), ""),
    }


def get_app_version(default: str = "0.1") -> str:
    """Return app version from environment.

    Supports both APP_VERSION and legacy lowercase version keys in .env.
    """
    load_env()
    version_value = _clean(os.getenv("APP_VERSION", os.getenv("version", default)), default)
    if not version_value:
        return default
    return version_value
