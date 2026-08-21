import datetime
import os
from zoneinfo import ZoneInfo

import requests


LOGIN_DATABASE_ENV_NAME = "DATABASE_spm"
LOGIN_RECORD_NAME = "Login"
LOGIN_WRONG_PASSWORD_PROPERTY_TYPES = frozenset({"rich_text", "select", "checkbox", "status"})
_login_property_schema: dict[str, tuple[str, str]] | None = None
_login_property_schema_database_id: str | None = None

try:
    from script.env_loader import load_env
except ModuleNotFoundError:
    import sys
    from pathlib import Path

    sys.path.append(str(Path(__file__).resolve().parent.parent))
    from script.env_loader import load_env


def _get_notion_settings() -> tuple[str, str]:
    load_env()

    token = os.getenv("NOTION_TOKEN", "").strip()
    database_id = os.getenv("DATABASE_ID", os.getenv("NOTION_DATABASE_ID", "")).strip()
    return token, database_id


def _get_login_notion_settings() -> tuple[str, str]:
    """Load the Notion credentials used by the SRCities login audit database."""
    load_env()
    token = os.getenv("NOTION_TOKEN", "").strip()
    database_id = os.getenv(LOGIN_DATABASE_ENV_NAME, "").strip()
    return token, database_id


def _rich_text_or_empty(value: str) -> list:
    text = (value or "").strip()
    if not text:
        return []

    # Notion rich_text text.content has a max length of 2000 characters.
    max_len = 2000
    chunks = [text[i:i + max_len] for i in range(0, len(text), max_len)]
    return [{"text": {"content": chunk}} for chunk in chunks]


def _wrong_password_property_value(wrong_password: bool, property_mode: str) -> dict:
    """Build a Yes/No value for common Notion property configurations."""
    label = "Yes" if wrong_password else "No"
    if property_mode == "checkbox":
        return {"checkbox": wrong_password}
    if property_mode == "select":
        return {"select": {"name": label}}
    if property_mode == "status":
        return {"status": {"name": label}}
    return {"rich_text": _rich_text_or_empty(label)}


def _normalized_property_name(name: str) -> str:
    """Normalize a Notion property name while preserving its configured key separately."""
    return " ".join(name.split()).casefold()


def _get_login_property_schema(
    database_id: str,
    headers: dict[str, str],
) -> dict[str, tuple[str, str]]:
    """Load and cache the actual property names and types for the login audit database."""
    global _login_property_schema, _login_property_schema_database_id

    if _login_property_schema is not None and _login_property_schema_database_id == database_id:
        return _login_property_schema

    response = requests.get(
        f"https://api.notion.com/v1/databases/{database_id}",
        headers=headers,
        timeout=10,
    )
    if response.status_code >= 300:
        raise RuntimeError("Notion login audit database could not be read.")

    payload = response.json()
    properties = payload.get("properties")
    if not isinstance(properties, dict):
        raise RuntimeError("Notion login audit database properties are unavailable.")

    schema: dict[str, tuple[str, str]] = {}
    for property_name, definition in properties.items():
        if not isinstance(property_name, str) or not isinstance(definition, dict):
            continue
        property_type = definition.get("type")
        if isinstance(property_type, str):
            schema[_normalized_property_name(property_name)] = (property_name, property_type)

    expected_property_types = {
        "name": "title",
        "date": "date",
        "ip address": "rich_text",
    }
    for normalized_name, expected_type in expected_property_types.items():
        configured_property = schema.get(normalized_name)
        if configured_property is None or configured_property[1] != expected_type:
            raise RuntimeError("Notion login audit database schema does not match the required properties.")

    wrong_password_property = schema.get("wrong pwd")
    if wrong_password_property is None or wrong_password_property[1] not in LOGIN_WRONG_PASSWORD_PROPERTY_TYPES:
        raise RuntimeError("Notion login audit database schema does not match the required properties.")

    _login_property_schema = schema
    _login_property_schema_database_id = database_id
    return schema


def _build_login_attempt_payload(
    ip_address: str,
    wrong_password: bool,
    timestamp: str,
    name: str,
    property_schema: dict[str, tuple[str, str]],
) -> dict:
    """Build one login-audit record for the Name, Date, IP address, and Wrong Pwd columns."""
    name_property, _ = property_schema["name"]
    date_property, _ = property_schema["date"]
    ip_address_property, _ = property_schema["ip address"]
    wrong_password_property, wrong_password_property_type = property_schema["wrong pwd"]
    return {
        "properties": {
            name_property: {"title": _rich_text_or_empty(name)},
            date_property: {"date": {"start": timestamp}},
            ip_address_property: {"rich_text": _rich_text_or_empty(ip_address)},
            wrong_password_property: _wrong_password_property_value(
                wrong_password,
                wrong_password_property_type,
            ),
        }
    }


def _build_notion_payload(
    question: str,
    ip: str,
    answer: str,
    app_name: str,
    token_input: int | None,
    token_output: int | None,
    token_mode: str,
) -> dict:
    properties = {
        "Title": {
            "title": [
                {
                    "text": {
                        "content": "chat log"
                    }
                }
            ]
        },
        "App name": {
            "rich_text": [
                {
                    "text": {
                        "content": app_name
                    }
                }
            ]
        },
        "questions": {
            "rich_text": _rich_text_or_empty(question)
        },
        "answer": {
            "rich_text": _rich_text_or_empty(answer)
        },
        "ip address": {
            "rich_text": [
                {
                    "text": {
                        "content": ip
                    }
                }
            ]
        },
        "datetime": {
            "date": {
                "start": datetime.datetime.now(ZoneInfo("Europe/Paris")).isoformat()
            }
        }
    }

    if token_mode == "number":
        properties["token_input"] = {"number": token_input}
        properties["token_output"] = {"number": token_output}
    else:
        token_input_text = "" if token_input is None else str(token_input)
        token_output_text = "" if token_output is None else str(token_output)
        properties["token_input"] = {"rich_text": _rich_text_or_empty(token_input_text)}
        properties["token_output"] = {"rich_text": _rich_text_or_empty(token_output_text)}

    return {
        "properties": properties
    }


def write_to_notion(
    question,
    ip,
    answer="",
    app_name="TSU_LLM_AIUseCase",
    token_input: int | None = None,
    token_output: int | None = None,
):
    notion_token, database_id = _get_notion_settings()
    if not notion_token or not database_id:
        raise ValueError("Missing Notion credentials. Set NOTION_TOKEN and DATABASE_ID in the environment.")

    url = "https://api.notion.com/v1/pages"

    headers = {
        "Authorization": f"Bearer {notion_token}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28"
    }

    data = {
        "parent": {"database_id": database_id},
        **_build_notion_payload(
            question,
            ip,
            answer,
            app_name,
            token_input,
            token_output,
            token_mode="number",
        ),
    }

    res = requests.post(url, json=data, headers=headers)
    if res.status_code >= 300:
        # Fallback in case token fields are configured as rich_text.
        fallback_data = {
            "parent": {"database_id": database_id},
            **_build_notion_payload(
                question,
                ip,
                answer,
                app_name,
                token_input,
                token_output,
                token_mode="rich_text",
            ),
        }
        fallback_res = requests.post(url, json=fallback_data, headers=headers)
        if fallback_res.status_code >= 300:
            raise RuntimeError(f"Notion write failed: {fallback_res.status_code} {fallback_res.text}")
        return

    return


def write_login_attempt_to_notion(
    ip_address: str,
    wrong_password: bool,
    name: str = LOGIN_RECORD_NAME,
) -> None:
    """Write one password-validation event to the SRCities Notion audit database.

    The attempted password is intentionally not accepted or stored. The
    ``Wrong Pwd`` property is written as Yes/No, with compatibility for rich
    text, select, checkbox, or status columns of that name.
    """
    notion_token, database_id = _get_login_notion_settings()
    if not notion_token or not database_id:
        raise ValueError(
            f"Missing Notion login-audit credentials. Set NOTION_TOKEN and {LOGIN_DATABASE_ENV_NAME}."
        )

    timestamp = datetime.datetime.now(ZoneInfo("Europe/Paris")).isoformat()
    headers = {
        "Authorization": f"Bearer {notion_token}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    }
    property_schema = _get_login_property_schema(database_id, headers)
    payload = {
        "parent": {"database_id": database_id},
        **_build_login_attempt_payload(
            ip_address=ip_address,
            wrong_password=wrong_password,
            timestamp=timestamp,
            name=name or LOGIN_RECORD_NAME,
            property_schema=property_schema,
        ),
    }
    response = requests.post(
        "https://api.notion.com/v1/pages",
        json=payload,
        headers=headers,
        timeout=10,
    )
    if response.status_code >= 300:
        raise RuntimeError("Notion login audit write failed.")


if __name__ == "__main__":
    write_to_notion("test", "127.0.0.1", "test answer")