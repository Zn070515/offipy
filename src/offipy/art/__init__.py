"""offipy 艺术层（v0.12）：演示文稿艺术分析与设计一致性系统。

只做建议（带 confidence 的可解释 Finding），默认不阻断；
strict 门禁仍归 0.11 几何审计层。
"""

from .adapters import MeasurementAdapter, PptxAuditAdapter, build_scene
from .analyze import analyze_deck, analyze_scene
from .compare import ArtReportDiff, compare_reports
from .features_registry import FEATURES, encode_features, feature_keys, feature_schema_version
from .feedback import (
    ART_FEEDBACK_FILE,
    DEFAULT_DIR,
    ArtFeedbackRecord,
    append,
    apply_feedback,
    load_records,
    recommend_adjustments,
)
from .merge import merge_scenes
from .models import (
    ART_REPORT_SCHEMA_VERSION,
    ART_SCHEMA_VERSION,
    ArtColor,
    ArtElement,
    ArtElementRef,
    ArtFinding,
    ArtReport,
    ArtScene,
    ArtSlide,
    ArtSlideReport,
    ArtTextRun,
    ArtWarning,
    DeckQualityReport,
    DimensionAssessment,
    ElementPixelEvidence,
    PixelColorShare,
    SlidePixelEvidence,
)
from .pixels import PixelEnricher
from .profiles import ArtProfile, get_profile, profile_names
from .render import render_html, render_markdown, report_to_json

__all__ = [
    "ART_FEEDBACK_FILE",
    "ART_REPORT_SCHEMA_VERSION",
    "ART_SCHEMA_VERSION",
    "DEFAULT_DIR",
    "FEATURES",
    "ArtColor",
    "ArtElement",
    "ArtElementRef",
    "ArtFeedbackRecord",
    "ArtFinding",
    "ArtProfile",
    "ArtReport",
    "ArtReportDiff",
    "ArtScene",
    "ArtSlide",
    "ArtSlideReport",
    "ArtTextRun",
    "ArtWarning",
    "DeckQualityReport",
    "DimensionAssessment",
    "ElementPixelEvidence",
    "MeasurementAdapter",
    "PixelColorShare",
    "PixelEnricher",
    "PptxAuditAdapter",
    "SlidePixelEvidence",
    "analyze_deck",
    "analyze_scene",
    "append",
    "apply_feedback",
    "build_scene",
    "compare_reports",
    "encode_features",
    "feature_keys",
    "feature_schema_version",
    "get_profile",
    "load_records",
    "merge_scenes",
    "profile_names",
    "recommend_adjustments",
    "render_html",
    "render_markdown",
    "report_to_json",
]
