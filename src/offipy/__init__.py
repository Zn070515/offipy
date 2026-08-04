from .core import (
    PROGIDS,
    connect,
    ensure_app,
    launch,
    quit_app,
    running,
)
from .exceptions import (
    ConversionError,
    OfficeUnavailableError,
    OffipyError,
    RemoteCallError,
    ServerStartError,
    UnsupportedPlatformError,
)

__version__ = "0.8.0"

__all__ = [
    "connect",
    "ensure_app",
    "launch",
    "quit_app",
    "running",
    "PROGIDS",
    "OffipyError",
    "OfficeUnavailableError",
    "ServerStartError",
    "RemoteCallError",
    "ConversionError",
    "UnsupportedPlatformError",
]
