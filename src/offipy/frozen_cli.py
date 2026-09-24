"""Frozen entrypoint for the offipy command-line interface."""

from __future__ import annotations

from offipy import cli
from offipy.runtime import configure_browser_env


def main(argv: list[str] | None = None) -> int | None:
    """Dispatch to the normal CLI without adding frozen-build behavior."""

    configure_browser_env()
    return cli.main(argv)


if __name__ == "__main__":
    raise SystemExit(main())
