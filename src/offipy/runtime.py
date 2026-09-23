"""Runtime and helper-process locations for development and packaged builds.

The Office-facing layers should not need to know whether offipy is running from
an interpreter checkout or from a frozen installation.  This module is the
single place that translates logical helpers (server, converter, Chromium,
and user data) into paths and commands.
"""

from __future__ import annotations

import sys
from pathlib import Path

SERVER_MODULE = "offipy.server"
SERVER_EXECUTABLE = "Offipy.Server.exe"
CONVERTER_EXECUTABLE = "Offipy.Converter.exe"
_CONVERTER_RELATIVE_PATH = Path("_vendor") / "html_to_editable_pptx" / "convert.py"
_CHROMIUM_DIRECTORY = "chromium"


def is_packaged() -> bool:
    """Return whether the current process is a frozen application build."""

    return bool(getattr(sys, "frozen", False))


def resource_root() -> Path:
    """Return the directory containing packaged resources or the source package."""

    if is_packaged():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def data_root() -> Path:
    """Return offipy's per-user writable data directory."""

    # Import lazily to keep this module dependency-free and preserve the
    # existing paths.py environment/platform selection in one place.
    from .paths import user_data_dir

    return user_data_dir()


def converter_script() -> Path:
    """Return the development converter script location."""

    return resource_root() / _CONVERTER_RELATIVE_PATH


def server_command(port: int) -> list[str]:
    """Build the command used to start the local Office server."""

    if is_packaged():
        return [str(resource_root() / SERVER_EXECUTABLE), "--port", str(port)]
    return [sys.executable, "-m", SERVER_MODULE, "--port", str(port)]


def converter_command(html: str | Path) -> list[str]:
    """Build the command used to convert one HTML deck."""

    if is_packaged():
        return [str(resource_root() / CONVERTER_EXECUTABLE), str(html)]
    return [sys.executable, str(converter_script()), str(html)]


def chromium_path() -> Path | None:
    """Return a bundled Chromium directory when one is present.

    Development builds continue to use Playwright's normal browser discovery;
    the packaged layout reserves a sibling ``chromium`` directory for the
    self-contained runtime introduced by the frozen-build stage.
    """

    if not is_packaged():
        return None
    path = resource_root() / _CHROMIUM_DIRECTORY
    return path if path.is_dir() else None
