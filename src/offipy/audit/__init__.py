"""PPTX 质量审计与几何回归：结构提取 + 静态审计 + 基线回归 + 报告 + 质量门禁。

`import offipy.audit` 不触发 python-pptx（惰性加载，见 audit/extract.py）。
"""

from .models import (
    ALL_RULE_IDS,
    AUDIT_SCHEMA_VERSION,
    RULE_AUTOFIT_GROW,
    RULE_AUTOFIT_SHRINK,
    RULE_BOUNDS_OFF_CANVAS,
    RULE_BOUNDS_PARTIAL,
    RULE_MARGIN_BOTTOM,
    RULE_MARGIN_LEFT,
    RULE_MARGIN_RIGHT,
    RULE_MARGIN_TOP,
    RULE_OVERLAP_COVERED_TEXT,
    RULE_OVERLAP_PARTIAL,
    RULE_TEXT_FIT_HORIZONTAL,
    RULE_TEXT_FIT_VERTICAL,
    AuditConfig,
    AuditFinding,
    AuditShapeRef,
    AuditWarning,
    JsonValue,
    PptxAuditReport,
    Severity,
    SuppressedFinding,
)

__all__ = [
    "AUDIT_SCHEMA_VERSION",
    "ALL_RULE_IDS",
    "RULE_AUTOFIT_GROW",
    "RULE_AUTOFIT_SHRINK",
    "RULE_BOUNDS_OFF_CANVAS",
    "RULE_BOUNDS_PARTIAL",
    "RULE_MARGIN_BOTTOM",
    "RULE_MARGIN_LEFT",
    "RULE_MARGIN_RIGHT",
    "RULE_MARGIN_TOP",
    "RULE_OVERLAP_COVERED_TEXT",
    "RULE_OVERLAP_PARTIAL",
    "RULE_TEXT_FIT_HORIZONTAL",
    "RULE_TEXT_FIT_VERTICAL",
    "AuditConfig",
    "AuditFinding",
    "AuditShapeRef",
    "AuditWarning",
    "JsonValue",
    "PptxAuditReport",
    "Severity",
    "SuppressedFinding",
]
