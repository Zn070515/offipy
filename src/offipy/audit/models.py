"""PPTX 质量审计：公共报告与配置模型（零第三方依赖）。

只用标准库 dataclasses/enum/typing。硬约束：`import offipy.audit`、
`from offipy.audit.models import Severity` 都不触发 python-pptx；
python-pptx 只在 audit/extract.py 与 audit/pptx.py 读文件时才惰性 import。

序列化：`to_dict()` 输出完全 JSON 安全（无 Enum/Path/set）；比较 Severity
必须按整数值（IntEnum），禁止字符串比较。
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from enum import IntEnum
from typing import Any, Literal

JsonValue = int | float | str | bool | None | list["JsonValue"] | dict[str, "JsonValue"]

# 报告 schema 版本：结构变更时递增（与 offipy 版本解耦，CI/下游可判稳定）
AUDIT_SCHEMA_VERSION = "0.1"

# ---------------------------------------------------------------- 严重度


# 序列化输出 "LOW"/"MID"/"HIGH"；比较必须按整数值（IntEnum），禁止字符串比较。
class Severity(IntEnum):
    LOW = 1
    MID = 2
    HIGH = 3


# ---------------------------------------------------------------- 稳定 rule_id

# 用户/CI 依赖这些 ID 而非自然语言 message；实现与测试必须引用本常量。
RULE_BOUNDS_PARTIAL = "geometry.bounds.partial"
RULE_BOUNDS_OFF_CANVAS = "geometry.bounds.off_canvas"
RULE_MARGIN_LEFT = "geometry.margin.left"
RULE_MARGIN_RIGHT = "geometry.margin.right"
RULE_MARGIN_TOP = "geometry.margin.top"
RULE_MARGIN_BOTTOM = "geometry.margin.bottom"
RULE_OVERLAP_PARTIAL = "geometry.overlap.partial"
RULE_OVERLAP_COVERED_TEXT = "geometry.overlap.covered_text"
RULE_TEXT_FIT_HORIZONTAL = "text.fit.horizontal"
RULE_TEXT_FIT_VERTICAL = "text.fit.vertical"
RULE_AUTOFIT_SHRINK = "text.autofit.shrink"
RULE_AUTOFIT_GROW = "text.autofit.grow"

ALL_RULE_IDS = (
    RULE_BOUNDS_PARTIAL,
    RULE_BOUNDS_OFF_CANVAS,
    RULE_MARGIN_LEFT,
    RULE_MARGIN_RIGHT,
    RULE_MARGIN_TOP,
    RULE_MARGIN_BOTTOM,
    RULE_OVERLAP_PARTIAL,
    RULE_OVERLAP_COVERED_TEXT,
    RULE_TEXT_FIT_HORIZONTAL,
    RULE_TEXT_FIT_VERTICAL,
    RULE_AUTOFIT_SHRINK,
    RULE_AUTOFIT_GROW,
)

FindingKind = Literal["bounds", "margin", "overlap", "text_fit", "autofit"]
SuppressionReason = Literal[
    "user_shape",
    "user_region",
    "page_number",
    "header_footer",
    "repeated_decoration",
    "full_bleed",
    "intentional_containment",
    "decorative_overlay",
    "text_on_background",
    "transparent_overlay",
    "decorative_layering",
]


# ---------------------------------------------------------------- 配置


@dataclass(frozen=True)
class AuditConfig:
    """审计配置。所有尺寸单位为英寸（in，幻灯片绝对坐标）。

    ignored_shapes: (slide_index 1-based, shape_id) 集合，用户显式豁免。
    ignored_regions: 英寸 (x, y, w, h) 列表，与 shape AABB 中心命中即豁免。
    """

    safe_margin_in: float = 0.2
    bounds_tolerance_in: float = 0.01
    ignore_full_bleed_shapes: bool = True
    ignore_repeated_decorations: bool = True
    ignore_page_numbers: bool = True
    ignore_headers_footers: bool = True
    ignored_shapes: set[tuple[int, int]] | None = None
    ignored_regions: list[tuple[float, float, float, float]] | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "safe_margin_in": self.safe_margin_in,
            "bounds_tolerance_in": self.bounds_tolerance_in,
            "ignore_full_bleed_shapes": self.ignore_full_bleed_shapes,
            "ignore_repeated_decorations": self.ignore_repeated_decorations,
            "ignore_page_numbers": self.ignore_page_numbers,
            "ignore_headers_footers": self.ignore_headers_footers,
            "ignored_shapes": (sorted(self.ignored_shapes) if self.ignored_shapes else None),
            "ignored_regions": self.ignored_regions,
        }


# ---------------------------------------------------------------- Finding 模型


@dataclass(frozen=True)
class AuditShapeRef:
    """Finding 引用的 Shape 轻量快照（slide_index 为 1-based）。"""

    slide_index: int
    shape_id: int
    name: str
    shape_type: str
    role: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "slide_index": self.slide_index,
            "shape_id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type,
            "role": self.role,
        }


@dataclass(frozen=True)
class AuditFinding:
    """单条审计发现。rule_id 稳定，message 是给用户看的中文描述。

    details 必须 JSON 安全（JsonValue）；confidence ∈ (0, 1]。
    """

    rule_id: str
    kind: FindingKind
    severity: Severity
    message: str
    primary: AuditShapeRef
    secondary: AuditShapeRef | None = None
    details: dict[str, JsonValue] = field(default_factory=dict)
    confidence: float = 1.0

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "rule_id": self.rule_id,
            "kind": self.kind,
            "severity": self.severity.name,
            "message": self.message,
            "primary": self.primary.to_dict(),
            "confidence": self.confidence,
        }
        if self.secondary is not None:
            d["secondary"] = self.secondary.to_dict()
        if self.details:
            d["details"] = self.details
        return d


@dataclass(frozen=True)
class SuppressedFinding:
    """被豁免的 Finding（用户 ignore 或自动角色识别）。不静默丢弃，全部可追溯。"""

    finding: AuditFinding
    reason: SuppressionReason

    def to_dict(self) -> dict[str, Any]:
        return {"finding": self.finding.to_dict(), "reason": self.reason}


@dataclass(frozen=True)
class AuditWarning:
    """解析异常 / 不支持结构 / 无法精确转换的几何，全部可追溯。"""

    slide_index: int | None
    shape_id: int | None
    code: str
    message: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "slide_index": self.slide_index,
            "shape_id": self.shape_id,
            "code": self.code,
            "message": self.message,
        }


# ---------------------------------------------------------------- 报告


@dataclass(frozen=True)
class SlideShapeSnapshot:
    """每页 Shape 几何快照，供 HTML/SVG 渲染（渲染视图，非 v0.12 ShapeInfo API）。

    尺寸单位为幻灯片绝对英寸；None 表示该维度缺失/无法解析。
    """

    slide_index: int  # 1-based
    shape_id: int
    name: str
    shape_type: str
    role: str
    left: float | None
    top: float | None
    width: float | None
    height: float | None
    z_order: int
    text: str
    is_rotated: bool
    geometry_unknown: bool

    def to_dict(self) -> dict[str, Any]:
        return {
            "slide_index": self.slide_index,
            "shape_id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type,
            "role": self.role,
            "left": self.left,
            "top": self.top,
            "width": self.width,
            "height": self.height,
            "z_order": self.z_order,
            "text": self.text,
            "is_rotated": self.is_rotated,
            "geometry_unknown": self.geometry_unknown,
        }


@dataclass
class PptxAuditReport:
    schema_version: str
    offipy_version: str
    path: str
    source_sha256: str
    slide_size: tuple[float, float]
    slide_count: int
    config: AuditConfig
    findings: list[AuditFinding] = field(default_factory=list)
    suppressed: list[SuppressedFinding] = field(default_factory=list)
    warnings: list[AuditWarning] = field(default_factory=list)
    shapes: list[SlideShapeSnapshot] = field(default_factory=list)
    # #128 富集：extract_presentation 的 _ShapeRecord 原样携带（不进 to_dict 序列化，
    # 仅供 PptxAuditAdapter 读取完整视觉证据；schemeClr 主题色不在此解析）
    records: list[Any] | None = field(default=None, repr=False)

    @property
    def max_severity(self) -> Severity | None:
        """最高严重度（无 finding 返回 None）。按整数值比较，禁止字符串比较。"""
        if not self.findings:
            return None
        return max(f.severity for f in self.findings)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "offipy_version": self.offipy_version,
            "path": self.path,
            "source_sha256": self.source_sha256,
            "slide_size": list(self.slide_size),
            "slide_count": self.slide_count,
            "config": self.config.to_dict(),
            "findings": [f.to_dict() for f in self.findings],
            "suppressed": [s.to_dict() for s in self.suppressed],
            "warnings": [w.to_dict() for w in self.warnings],
            "shapes": [s.to_dict() for s in self.shapes],
            "max_severity": (self.max_severity.name if self.max_severity is not None else None),
        }

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    def to_markdown(self) -> str:
        from .render import render_markdown

        return render_markdown(self)

    def to_html(self, *, slides_dir: str | None = None) -> str:
        from .render import render_html

        return render_html(self, slides_dir=slides_dir)


# ---------------------------------------------------------------- 基线回归


@dataclass(frozen=True)
class DiffShapeChange:
    """基线回归中单个形状变化（新增/删除/移动/缩放/文本）。

    details 记录旧/新几何与文本（英寸），完全 JSON 安全。
    """

    kind: Literal["added", "removed", "moved", "resized", "text"]
    slide_index: int
    shape_id: int
    name: str
    shape_type: str
    details: dict[str, JsonValue] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "kind": self.kind,
            "slide_index": self.slide_index,
            "shape_id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type,
        }
        if self.details:
            d["details"] = self.details
        return d


@dataclass(frozen=True)
class ChangedFinding:
    """同一 Finding（形状已匹配）的严重度变化。worsened=True 表示候选严重度更高。"""

    rule_id: str
    kind: FindingKind
    old_severity: Severity
    new_severity: Severity
    primary: AuditShapeRef
    secondary: AuditShapeRef | None = None
    details: dict[str, JsonValue] = field(default_factory=dict)

    @property
    def worsened(self) -> bool:
        """严重度是否上升（按整数值比较，禁止字符串比较）。"""
        return self.new_severity > self.old_severity

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "rule_id": self.rule_id,
            "kind": self.kind,
            "old_severity": self.old_severity.name,
            "new_severity": self.new_severity.name,
            "worsened": self.worsened,
            "primary": self.primary.to_dict(),
        }
        if self.secondary is not None:
            d["secondary"] = self.secondary.to_dict()
        if self.details:
            d["details"] = self.details
        return d


@dataclass
class PptxDiffReport:
    """基线 vs 候选 PPTX 回归报告：形状增删/移动/缩放/文本 + Finding 新增/已解决/变化。

    max_severity 语义：候选相对基线**新增或恶化**的最高严重度
    （`--fail-on-new` 门槛的依据）；无新增/恶化返回 None。
    """

    schema_version: str
    offipy_version: str
    baseline_path: str
    candidate_path: str
    baseline_sha256: str
    candidate_sha256: str
    baseline_slide_count: int
    candidate_slide_count: int
    baseline_findings: list[AuditFinding] = field(default_factory=list)
    candidate_findings: list[AuditFinding] = field(default_factory=list)
    added_findings: list[AuditFinding] = field(default_factory=list)
    resolved_findings: list[AuditFinding] = field(default_factory=list)
    changed_findings: list[ChangedFinding] = field(default_factory=list)
    added_shapes: list[DiffShapeChange] = field(default_factory=list)
    removed_shapes: list[DiffShapeChange] = field(default_factory=list)
    moved_shapes: list[DiffShapeChange] = field(default_factory=list)
    resized_shapes: list[DiffShapeChange] = field(default_factory=list)
    text_changes: list[DiffShapeChange] = field(default_factory=list)
    unmatched_baseline: list[AuditShapeRef] = field(default_factory=list)
    unmatched_candidate: list[AuditShapeRef] = field(default_factory=list)
    warnings: list[AuditWarning] = field(default_factory=list)

    @property
    def added_slides(self) -> int:
        return max(0, self.candidate_slide_count - self.baseline_slide_count)

    @property
    def removed_slides(self) -> int:
        return max(0, self.baseline_slide_count - self.candidate_slide_count)

    @property
    def new_or_worsened(self) -> list[AuditFinding | ChangedFinding]:
        """候选新增或恶化的 Finding（`--fail-on-new` 只看这些）。"""
        return [*self.added_findings, *(c for c in self.changed_findings if c.worsened)]

    def gate_severity(self) -> Severity | None:
        """候选新增或恶化的最高严重度；无则 None（门槛不触发）。"""
        sevs: list[int] = [f.severity for f in self.added_findings]
        sevs += [c.new_severity for c in self.changed_findings if c.worsened]
        return Severity(max(sevs)) if sevs else None

    def to_dict(self) -> dict[str, Any]:
        gate = self.gate_severity()
        return {
            "schema_version": self.schema_version,
            "offipy_version": self.offipy_version,
            "baseline_path": self.baseline_path,
            "candidate_path": self.candidate_path,
            "baseline_sha256": self.baseline_sha256,
            "candidate_sha256": self.candidate_sha256,
            "baseline_slide_count": self.baseline_slide_count,
            "candidate_slide_count": self.candidate_slide_count,
            "added_slides": self.added_slides,
            "removed_slides": self.removed_slides,
            "baseline_findings": [f.to_dict() for f in self.baseline_findings],
            "candidate_findings": [f.to_dict() for f in self.candidate_findings],
            "added_findings": [f.to_dict() for f in self.added_findings],
            "resolved_findings": [f.to_dict() for f in self.resolved_findings],
            "changed_findings": [c.to_dict() for c in self.changed_findings],
            "added_shapes": [s.to_dict() for s in self.added_shapes],
            "removed_shapes": [s.to_dict() for s in self.removed_shapes],
            "moved_shapes": [s.to_dict() for s in self.moved_shapes],
            "resized_shapes": [s.to_dict() for s in self.resized_shapes],
            "text_changes": [s.to_dict() for s in self.text_changes],
            "unmatched_baseline": [s.to_dict() for s in self.unmatched_baseline],
            "unmatched_candidate": [s.to_dict() for s in self.unmatched_candidate],
            "warnings": [w.to_dict() for w in self.warnings],
            "max_new_severity": (gate.name if gate is not None else None),
        }

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    def to_markdown(self) -> str:
        from .render import render_markdown

        return render_markdown(self)

    def to_html(self, *, slides_dir: str | None = None) -> str:
        from .render import render_html

        return render_html(self, slides_dir=slides_dir)
