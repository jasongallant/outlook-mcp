"""Outlook/M365 MCP server entrypoint (stdio transport)."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from .tools import register_tools


def build_server() -> FastMCP:
    mcp = FastMCP("outlook-mcp")
    register_tools(mcp)
    return mcp


mcp = build_server()


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
