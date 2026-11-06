"""Settings loaded from ~/.config/outlook-mcp/settings.json.

Create the file once to configure your M365 tenant and preferences:

    mkdir -p ~/.config/outlook-mcp
    cp settings.json.example ~/.config/outlook-mcp/settings.json

All keys are optional — defaults apply when missing.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

CONFIG_DIR = Path.home() / ".config" / "outlook-mcp"
CONFIG_PATH = CONFIG_DIR / "settings.json"

_DEFAULTS: dict = {
    "tenant": "common",
    "user_email": "",
    "timezone": "America/New_York",
    "writes_enabled": False,
    "mail_writes_enabled": False,
    "calendar_writes_enabled": False,
    "todo_writes_enabled": False,
}

_TRUTHY = {"1", "true", "yes", "on"}


def _load() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def get(key: str):
    return _load().get(key, _DEFAULTS.get(key))


def get_tenant() -> str:
    return str(get("tenant") or _DEFAULTS["tenant"])


def get_user_email() -> str:
    return str(get("user_email") or "")


def get_timezone() -> str:
    return str(get("timezone") or _DEFAULTS["timezone"])


def _flag(env_name: str, setting_key: str) -> bool:
    """A boolean toggle: env var wins if set, else the settings.json key."""
    env = os.environ.get(env_name, "").strip().lower()
    if env:
        return env in _TRUTHY
    return bool(get(setting_key))


def writes_enabled() -> bool:
    """Master write switch: enables all write groups (mail, calendar, To Do)."""
    return _flag("OUTLOOK_MCP_WRITES", "writes_enabled")


def mail_writes_enabled() -> bool:
    """Email write tools (send/reply/mark-read/delete). On if master or mail flag."""
    return writes_enabled() or _flag("OUTLOOK_MCP_MAIL_WRITES", "mail_writes_enabled")


def calendar_writes_enabled() -> bool:
    """Calendar write tools (create/update/respond). On if master or calendar flag."""
    return writes_enabled() or _flag("OUTLOOK_MCP_CALENDAR_WRITES", "calendar_writes_enabled")


def todo_writes_enabled() -> bool:
    """To Do write tools (create/update/complete/delete task, create list). On if master or todo flag."""
    return writes_enabled() or _flag("OUTLOOK_MCP_TODO_WRITES", "todo_writes_enabled")


def save(updates: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    current = _load()
    current.update(updates)
    CONFIG_PATH.write_text(json.dumps(current, indent=2) + "\n")
