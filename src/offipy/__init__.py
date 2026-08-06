from . import direct
from .api import Excel, Ppt, RemoteExcel, RemotePpt, RemoteWord, Word, op
from .audit import (
    AuditConfig,
    AuditFinding,
    PptxAuditReport,
    PptxDiffReport,
    Severity,
    audit_pptx,
    compare_pptx,
)
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
from .models import PLACEHOLDER_TYPE_NAMES, SlideTextRecord

__version__ = "0.11.1"

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
    "SlideTextRecord",
    "PLACEHOLDER_TYPE_NAMES",
    "Severity",
    "AuditConfig",
    "AuditFinding",
    "PptxAuditReport",
    "PptxDiffReport",
    "audit_pptx",
    "compare_pptx",
]
