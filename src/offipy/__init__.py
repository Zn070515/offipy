from typing import Any

from . import direct
from .api import Excel, Ppt, RemoteExcel, RemotePpt, RemoteWord, Word, op
from .art import (
    ART_REPORT_SCHEMA_VERSION,
    ART_SCHEMA_VERSION,
    ArtColor,
    ArtElement,
    ArtElementRef,
    ArtFinding,
    ArtProfile,
    ArtReport,
    ArtReportDiff,
    ArtScene,
    ArtSlide,
    ArtSlideReport,
    ArtTextRun,
    ArtWarning,
    DeckQualityReport,
    DimensionAssessment,
    ElementPixelEvidence,
    PixelColorShare,
    PixelEnricher,
    SlidePixelEvidence,
    analyze_deck,
    analyze_scene,
    build_scene,
    compare_reports,
    get_profile,
    merge_scenes,
    profile_names,
    render_html,
    render_markdown,
    report_to_json,
)
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
from .models import (
    PLACEHOLDER_TYPE_NAMES,
    SHAPE_TYPE_NAMES,
    ShapeInfo,
    SlideTextRecord,
    shape_type_name,
)

__version__ = "0.20.0"

__all__ = [
    "ART_REPORT_SCHEMA_VERSION",
    "ART_SCHEMA_VERSION",
    "PLACEHOLDER_TYPE_NAMES",
    "PROGIDS",
    "SHAPE_TYPE_NAMES",
    "AnimationSpec",
    "ArtColor",
    "ArtElement",
    "ArtElementRef",
    "ArtFinding",
    "ArtProfile",
    "ArtReport",
    "ArtReportDiff",
    "ArtScene",
    "ArtSlide",
    "ArtSlideReport",
    "ArtTextRun",
    "ArtWarning",
    "AuditConfig",
    "AuditFinding",
    "ComOperationError",
    "ConversionError",
    "DeckQualityReport",
    "DimensionAssessment",
    "ElementPixelEvidence",
    "Excel",
    "FileConflictError",
    "InvalidArgumentError",
    "OfficeUnavailableError",
    "OffipyError",
    "PixelColorShare",
    "PixelEnricher",
    "Ppt",
    "PptxAuditReport",
    "PptxDiffReport",
    "ProtocolError",
    "RemoteCallError",
    "RemoteExcel",
    "RemotePpt",
    "RemoteWord",
    "ServerStartError",
    "Severity",
    "ShapeInfo",
    "SlidePixelEvidence",
    "SlideTextRecord",
    "TargetNotFoundError",
    "TransitionSpec",
    "UnsupportedPlatformError",
    "Word",
    "analyze_deck",
    "analyze_scene",
    "apply_animations",
    "apply_transitions",
    "audit_pptx",
    "build_scene",
    "compare_pptx",
    "compare_reports",
    "connect",
    "direct",
    "ensure_app",
    "get_profile",
    "launch",
    "merge_scenes",
    "op",
    "profile_names",
    "quit_app",
    "render_html",
    "render_markdown",
    "report_to_json",
    "running",
    "shape_type_name",
]

# animations 符号惰性导出：offipy.animations 包 __init__ 经 .apply 拉 python-pptx，
# 而核心 `import offipy` 承诺零额外依赖（deck extra 才装 python-pptx/lxml）。PEP 562
# __getattr__ 把 4 个符号的加载推迟到真正访问时，保住 test_import_offipy_does_not_load_pptx
# 等「import offipy 不加载 pptx」契约。
_ANIMATION_LAZY_EXPORTS = frozenset(
    {"AnimationSpec", "TransitionSpec", "apply_animations", "apply_transitions"}
)


def __getattr__(name: str) -> Any:
    """PEP 562 惰性导出 animations 符号（访问时才触发 offipy.animations 包）。"""
    if name not in _ANIMATION_LAZY_EXPORTS:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    from .animations import (
        AnimationSpec,
        TransitionSpec,
        apply_animations,
        apply_transitions,
    )

    return {
        "AnimationSpec": AnimationSpec,
        "TransitionSpec": TransitionSpec,
        "apply_animations": apply_animations,
        "apply_transitions": apply_transitions,
    }[name]
