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
}


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


def writes_enabled() -> bool:
    """True if OUTLOOK_MCP_WRITES env var is truthy, or settings.json writes_enabled is true."""
    env = os.environ.get("OUTLOOK_MCP_WRITES", "").strip().lower()
    if env:
        return env in {"1", "true", "yes", "on"}
    return bool(get("writes_enabled"))


def save(updates: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    current = _load()
    current.update(updates)
    CONFIG_PATH.write_text(json.dumps(current, indent=2) + "\n")
