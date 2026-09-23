"""Frozen entrypoint for the resident Office server."""

from __future__ import annotations

from offipy import server


def main(argv: list[str] | None = None) -> None:
    """Dispatch to the resident server argument parser and lifecycle."""

    server.main(argv)


if __name__ == "__main__":
    main()
