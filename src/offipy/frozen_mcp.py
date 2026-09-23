"""Frozen entrypoint for the MCP stdio server."""

from __future__ import annotations

from offipy import mcp_server


def main() -> None:
    """Dispatch to the normal MCP server."""

    mcp_server.main()


if __name__ == "__main__":
    main()
