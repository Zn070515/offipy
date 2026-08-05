from . import direct
from .api import Excel, Ppt, RemoteExcel, RemotePpt, RemoteWord, Word, op
from .core import (
    PROGIDS,
    connect,
    ensure_app,
    launch,
    quit_app,
    running,
)
from .exceptions import (
    ComOperationError,
    ConversionError,
    FileConflictError,
    InvalidArgumentError,
    OfficeUnavailableError,
    OffipyError,
    ProtocolError,
    RemoteCallError,
    ServerStartError,
    TargetNotFoundError,
    UnsupportedPlatformError,
)

__version__ = "0.9.0a1"

__all__ = [
    "Excel",
    "Word",
    "Ppt",
    "RemoteExcel",
    "RemoteWord",
    "RemotePpt",
    "direct",
    "op",
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
    "TargetNotFoundError",
    "ConversionError",
    "UnsupportedPlatformError",
    "InvalidArgumentError",
    "FileConflictError",
    "ComOperationError",
    "ProtocolError",
]
