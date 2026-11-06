"""Field lists, tool registration, and thread offload. No network."""

from __future__ import annotations

import asyncio
import threading

import pytest
from exchangelib import CalendarItem, Contact, Message, Task
from exchangelib.version import EXCHANGE_O365, Version
from mcp.server.fastmcp import FastMCP

from outlook_mcp import tools


@pytest.mark.parametrize(
    ("model", "fields"),
    [
        (Message, tools._MSG_FIELDS),
        (Message, tools._MSG_FULL_FIELDS),
        (CalendarItem, tools._EVENT_FIELDS),
        (Task, tools._TASK_FIELDS),
        (Task, tools._TASK_FULL_FIELDS),
        (Contact, tools._CONTACT_FIELDS),
    ],
)
def test_field_names_exist_on_model(model, fields):
    # .only() raises on an unknown name, so a typo here breaks a tool at call time.
    supported = {f.name for f in model.supported_fields(version=Version(build=EXCHANGE_O365))}
    assert set(fields) <= supported


@pytest.mark.parametrize("fields", [tools._MSG_FIELDS, tools._MSG_FULL_FIELDS])
def test_message_fields_never_request_mime(fields):
    assert "mime_content" not in fields


def test_threaded_tool_runs_off_event_loop_thread():
    server = FastMCP("t")
    tool = tools._threaded_tool(server)

    @tool(description="d")
    def where(x: int, y: str = "a") -> dict:
        return {"thread": threading.current_thread().name, "x": x, "y": y}

    async def run():
        loop_thread = threading.current_thread().name
        out = await server._tool_manager.call_tool("where", {"x": 1}, convert_result=False)
        return loop_thread, out

    loop_thread, out = asyncio.run(run())
    assert out["thread"] != loop_thread
    assert (out["x"], out["y"]) == (1, "a")


def test_threaded_tool_keeps_schema():
    def plain(x: int, y: str = "a", z: list[str] | None = None) -> dict:
        return {}

    direct, wrapped = FastMCP("direct"), FastMCP("wrapped")
    direct.tool(description="d")(plain)
    tools._threaded_tool(wrapped)(description="d")(plain)
    a = asyncio.run(direct.list_tools())[0].model_dump(mode="json")
    b = asyncio.run(wrapped.list_tools())[0].model_dump(mode="json")
    assert a == b


def test_all_tools_register_with_writes_enabled(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_WRITES", "1")
    server = FastMCP("t")
    tools.register_tools(server)
    names = {t.name for t in asyncio.run(server.list_tools())}
    assert {"get_attachment", "list_attachments", "send_email", "reply_to_message"} <= names


def _registered(monkeypatch, **flags):
    for name in ("WRITES", "MAIL_WRITES", "CALENDAR_WRITES", "TODO_WRITES"):
        monkeypatch.setenv(f"OUTLOOK_MCP_{name}", flags.get(name, "0"))
    server = FastMCP("t")
    tools.register_tools(server)
    return {t.name: t for t in asyncio.run(server.list_tools())}


def test_get_attachment_is_available_without_writes(monkeypatch):
    names = _registered(monkeypatch)
    assert "get_attachment" in names
    assert "send_email" not in names


def test_todo_flag_alone_keeps_email_sending_off(monkeypatch):
    names = _registered(monkeypatch, TODO_WRITES="1")
    assert "create_task" in names
    assert "get_attachment" in names
    assert "send_email" not in names
    assert "reply_to_message" not in names
    assert "create_event" not in names


def test_mail_flag_alone_exposes_send_with_attachments(monkeypatch):
    names = _registered(monkeypatch, MAIL_WRITES="1")
    assert "attachments" in names["send_email"].inputSchema["properties"]
    assert "attachments" in names["reply_to_message"].inputSchema["properties"]
    assert "create_task" not in names
    assert "create_event" not in names
