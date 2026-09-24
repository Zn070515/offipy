"""Frozen entrypoint for the MCP stdio server."""

from __future__ import annotations

from offipy import mcp_server
from offipy.runtime import configure_browser_env


def main() -> None:
    """Dispatch to the normal MCP server."""

    configure_browser_env()
    mcp_server.main()


if __name__ == "__main__":
    main()
