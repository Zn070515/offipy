"""art 数据模型（v0.12 契约冻结，rev2.1 一次定死全部字段）。

公开 slide 索引一律 1-based；坐标/尺寸按页宽高归一化但不限 [0,1]；
字号规范：font_size + font_size_unit + font_size_norm（普通字段，适配器直接填），
规则只消费 font_size_norm，模型不做任何 norm 回填。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal

if TYPE_CHECKING:
    from offipy.audit import Severity

ART_SCHEMA_VERSION = "0.2"
ART_REPORT_SCHEMA_VERSION = "0.4"

Grade = Literal["excellent", "good", "attention", "poor"]
AssessmentStatus = Literal["assessed", "insufficient_evidence", "not_applicable"]

JsonValue = int | float | str | bool | None | list["JsonValue"] | dict[str, "JsonValue"]


@dataclass(frozen=True)
class ArtColor:
    r: int
    g: int
    b: int
    a: float = 1.0

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ArtColor | None:
        """未受信 color dict：r/g/b 缺失或非数字 → None（无颜色证据），不抛 ValueError。"""
        try:
            r, g, b = (int(v) for v in (data["r"], data["g"], data["b"]))
            a = float(data.get("a", 1.0))
        except (ValueError, TypeError, KeyError):
            return None
        return cls(r, g, b, a)

    def with_alpha_over(self, other: ArtColor) -> ArtColor:
        """self 盖在 other 之上的结果色（标准 alpha 合成）。"""
        sa = self.a
        da = other.a
        out_a = sa + da * (1 - sa)
        if out_a == 0:
            return ArtColor(0, 0, 0, 0.0)
        out_r = round((self.r * sa + other.r * da * (1 - sa)) / out_a)
        out_g = round((self.g * sa + other.g * da * (1 - sa)) / out_a)
        out_b = round((self.b * sa + other.b * da * (1 - sa)) / out_a)
        return ArtColor(out_r, out_g, out_b, round(out_a, 6))

    def to_dict(self) -> dict[str, Any]:
        return {"r": self.r, "g": self.g, "b": self.b, "a": self.a}


@dataclass(frozen=True)
class ArtTextRun:
    text: str
    font_size: float | None = None
    font_size_unit: str = "unknown"
    font_family: str | None = None
    color: ArtColor | None = None  # 该 run 的前景色（文本前景）

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "font_size": self.font_size,
            "font_size_unit": self.font_size_unit,
            "font_family": self.font_family,
            "color": self.color.to_dict() if self.color else None,
        }


@dataclass(frozen=True)
class ArtElementRef:
    slide_index: int
    element_id: str
    kind: str
    role: str

    def to_dict(self) -> dict[str, str | int]:
        return {
            "slide_index": self.slide_index,
            "element_id": self.element_id,
            "kind": self.kind,
            "role": self.role,
        }


@dataclass(frozen=True)
class ArtElement:
    element_id: str
    kind: str  # "text" | "image" | "shape"
    role: str  # "title" | "body" | "image" | "background" | ...
    x: float
    y: float
    width: float
    height: float
    slide_index: int
    foreground: ArtColor | None = None
    background: ArtColor | None = None
    border: ArtColor | None = None
    is_background: bool = False
    text: str = ""
    font_size: float | None = None
    font_size_unit: str = "unknown"
    font_size_norm: float | None = None  # 普通字段：适配器填，规则只读
    runs: list[ArtTextRun] = field(default_factory=list)
    natural_width: float | None = None
    natural_height: float | None = None
    source: str = "measurement"  # "measurement" | "pptx" | "merged"
    evidence: dict[str, Any] = field(default_factory=dict)
    container: bool = False
    decoration: bool = False
    pixel_evidence: ElementPixelEvidence | None = None
    opacity: float | None = None  # 元素级透明度 0-1；None=无证据（Task 4 消费）
    decoded_width: float | None = None  # 图片解码宽度（naturalWidth），漂移计算用
    decoded_height: float | None = None
    fill_kind: str | None = None  # "solid"|"gradient"|"shadow"|"image"|None（Task 5 消费）

    @property
    def area(self) -> float:
        return self.width * self.height

    def has_text(self) -> bool:
        return bool(self.text and self.text.strip())

    def to_dict(self) -> dict[str, Any]:
        return {
            "element_id": self.element_id,
            "kind": self.kind,
            "role": self.role,
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "slide_index": self.slide_index,
            "foreground": self.foreground.to_dict() if self.foreground else None,
            "background": self.background.to_dict() if self.background else None,
            "border": self.border.to_dict() if self.border else None,
            "is_background": self.is_background,
            "text": self.text,
            "font_size": self.font_size,
            "font_size_unit": self.font_size_unit,
            "font_size_norm": self.font_size_norm,
            "runs": [r.to_dict() for r in self.runs],
            "natural_width": self.natural_width,
            "natural_height": self.natural_height,
            "source": self.source,
            "evidence": self.evidence,
            "container": self.container,
            "decoration": self.decoration,
            "pixel_evidence": self.pixel_evidence.to_dict() if self.pixel_evidence else None,
            "opacity": self.opacity,
            "decoded_width": self.decoded_width,
            "decoded_height": self.decoded_height,
            "fill_kind": self.fill_kind,
        }


@dataclass
class ArtSlide:
    index: int  # 1-based
    width: float
    height: float
    elements: list[ArtElement] = field(default_factory=list)
    background_color: ArtColor | None = None
    pixel_evidence: SlidePixelEvidence | None = None

    def by_id(self, element_id: str) -> ArtElement | None:
        return next((e for e in self.elements if e.element_id == element_id), None)


@dataclass
class ArtScene:
    slides: list[ArtSlide] = field(default_factory=list)
    width_unit: str = "px"  # "px" | "pt"
    warnings: list[ArtWarning] = field(default_factory=list)
    sources: set[str] = field(default_factory=set)
    metadata: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ArtScene:
        """手写/测试场景输入协议（与 MeasurementAdapter 的真实格式分离）。"""
        slides: list[ArtSlide] = []
        for i, s in enumerate(data.get("slides", [])):
            idx = int(s.get("index") or i + 1)
            height = float(s.get("height", 1080.0))
            elements = [_element_from_dict(e, idx, height) for e in s.get("elements", [])]
            bg = s.get("background_color")
            pe = s.get("pixel_evidence")
            slides.append(
                ArtSlide(
                    index=idx,
                    width=float(s.get("width", 1920.0)),
                    height=height,
                    elements=elements,
                    background_color=ArtColor.from_dict(bg) if bg else None,
                    pixel_evidence=SlidePixelEvidence.from_dict(pe) if pe else None,
                )
            )
        return cls(
            slides=slides,
            width_unit=data.get("width_unit", "px"),
            warnings=[ArtWarning(w["code"], w["message"]) for w in data.get("warnings", [])],
            sources=set(data.get("sources", [])),
            metadata=data.get("metadata", {}).copy(),
        )

    def by_slide(self, index: int) -> ArtSlide | None:
        return next((s for s in self.slides if s.index == index), None)


def _element_from_dict(data: dict[str, Any], slide_index: int, height: float) -> ArtElement:
    fg = data.get("foreground") or data.get("color")
    bg = data.get("background")
    pe = data.get("pixel_evidence")
    if isinstance(bg, bool):  # 旧格式 bool background 标志 → is_background
        return ArtElement(
            element_id=data["element_id"],
            kind=data.get("kind", "shape"),
            role=data.get("role", "body"),
            x=float(data.get("x", 0.0)),
            y=float(data.get("y", 0.0)),
            width=float(data.get("width", 0.0)),
            height=float(data.get("height", 0.0)),
            slide_index=slide_index,
            foreground=_color_from(fg),
            is_background=bg,
            text=data.get("text", ""),
            font_size=float(data["font_size"]) if data.get("font_size") else None,
            font_size_unit=data.get("font_size_unit", "unknown"),
            font_size_norm=_norm_of(data, height),
            runs=[
                ArtTextRun(
                    text=rt.get("text", ""),
                    font_size=float(rt["font_size"]) if rt.get("font_size") else None,
                    font_size_unit=rt.get("font_size_unit", "unknown"),
                    font_family=rt.get("font_family"),
                    color=_color_from(rt.get("color")),
                )
                for rt in data.get("runs", [])
            ],
            natural_width=float(data["natural_width"]) if data.get("natural_width") else None,
            natural_height=float(data["natural_height"]) if data.get("natural_height") else None,
            container=bool(data.get("container")),
            decoration=bool(data.get("decoration")),
            pixel_evidence=ElementPixelEvidence.from_dict(pe) if pe else None,
            opacity=_opt_float(data.get("opacity")),
            decoded_width=_opt_float(data.get("decoded_width")),
            decoded_height=_opt_float(data.get("decoded_height")),
            fill_kind=data.get("fill_kind"),
        )
    return ArtElement(
        element_id=data["element_id"],
        kind=data.get("kind", "shape"),
        role=data.get("role", "body"),
        x=float(data.get("x", 0.0)),
        y=float(data.get("y", 0.0)),
        width=float(data.get("width", 0.0)),
        height=float(data.get("height", 0.0)),
        slide_index=slide_index,
        foreground=_color_from(fg),
        background=_color_from(bg),
        border=_color_from(data.get("border")),
        is_background=bool(data.get("is_background")),
        text=data.get("text", ""),
        font_size=float(data["font_size"]) if data.get("font_size") else None,
        font_size_unit=data.get("font_size_unit", "unknown"),
        font_size_norm=_norm_of(data, height),
        runs=[
            ArtTextRun(
                text=rt.get("text", ""),
                font_size=float(rt["font_size"]) if rt.get("font_size") else None,
                font_size_unit=rt.get("font_size_unit", "unknown"),
                font_family=rt.get("font_family"),
                color=_color_from(rt.get("color")),
            )
            for rt in data.get("runs", [])
        ],
        natural_width=float(data["natural_width"]) if data.get("natural_width") else None,
        natural_height=float(data["natural_height"]) if data.get("natural_height") else None,
        container=bool(data.get("container")),
        decoration=bool(data.get("decoration")),
        pixel_evidence=ElementPixelEvidence.from_dict(pe) if pe else None,
        opacity=_opt_float(data.get("opacity")),
        decoded_width=_opt_float(data.get("decoded_width")),
        decoded_height=_opt_float(data.get("decoded_height")),
        fill_kind=data.get("fill_kind"),
    )


def _norm_of(data: dict[str, Any], height: float) -> float | None:
    """手写场景的 norm：显式提供用显式值；否则 px/pt 同单位直接除。"""
    explicit = data.get("font_size_norm")
    if explicit is not None:
        return float(explicit)
    fs = data.get("font_size")
    unit = data.get("font_size_unit", "unknown")
    if fs is None or unit == "unknown" or not height:
        return None
    return float(fs) / height


def _color_from(c: object) -> ArtColor | None:
    if c is None:
        return None
    if isinstance(c, ArtColor):
        return c
    if isinstance(c, dict):
        return ArtColor.from_dict(c)
    if isinstance(c, (list, tuple)) and len(c) >= 3:
        return ArtColor(int(c[0]), int(c[1]), int(c[2]), float(c[3]) if len(c) > 3 else 1.0)
    return None


def _opt_float(v: Any) -> float | None:
    """序列化可选浮点：None 原样，其余转 float。"""
    return float(v) if v is not None else None


PixelEvidenceMethod = Literal[
    "declared_verified",
    "declared_not_found",
    "center_fill_verified",
    "complex_background",
    "unsupported",
]


@dataclass(frozen=True)
class PixelColorShare:
    color: ArtColor
    ratio: float

    def to_dict(self) -> dict[str, Any]:
        return {"color": self.color.to_dict(), "ratio": self.ratio}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PixelColorShare:
        color = ArtColor.from_dict(data["color"])
        if color is None:
            color = ArtColor(0, 0, 0)  # 损坏颜色兜底，不崩
        return cls(color, float(data["ratio"]))


@dataclass(frozen=True)
class SlidePixelEvidence:
    background: ArtColor | None = None
    background_confidence: float | None = None
    background_uniformity: float | None = None
    palette: list[PixelColorShare] = field(default_factory=list)
    background_like_ratio: float | None = None
    method: PixelEvidenceMethod = "unsupported"
    unsupported_reason: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "background": self.background.to_dict() if self.background else None,
            "background_confidence": self.background_confidence,
            "background_uniformity": self.background_uniformity,
            "palette": [p.to_dict() for p in self.palette],
            "background_like_ratio": self.background_like_ratio,
            "method": self.method,
            "unsupported_reason": self.unsupported_reason,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> SlidePixelEvidence:
        return cls(
            background=ArtColor.from_dict(data["background"]) if data.get("background") else None,
            background_confidence=_opt_float(data.get("background_confidence")),
            background_uniformity=_opt_float(data.get("background_uniformity")),
            palette=[PixelColorShare.from_dict(p) for p in data.get("palette", [])],
            background_like_ratio=_opt_float(data.get("background_like_ratio")),
            method=data.get("method", "unsupported"),
            unsupported_reason=data.get("unsupported_reason"),
        )


@dataclass(frozen=True)
class ElementPixelEvidence:
    foreground: ArtColor | None = None
    background: ArtColor | None = None
    foreground_match_ratio: float | None = None
    background_match_ratio: float | None = None
    background_complexity: float | None = None
    color_confidence: float = 0.0
    method: PixelEvidenceMethod = "unsupported"
    unsupported_reason: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "foreground": self.foreground.to_dict() if self.foreground else None,
            "background": self.background.to_dict() if self.background else None,
            "foreground_match_ratio": self.foreground_match_ratio,
            "background_match_ratio": self.background_match_ratio,
            "background_complexity": self.background_complexity,
            "color_confidence": self.color_confidence,
            "method": self.method,
            "unsupported_reason": self.unsupported_reason,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ElementPixelEvidence:
        return cls(
            foreground=ArtColor.from_dict(data["foreground"]) if data.get("foreground") else None,
            background=ArtColor.from_dict(data["background"]) if data.get("background") else None,
            foreground_match_ratio=_opt_float(data.get("foreground_match_ratio")),
            background_match_ratio=_opt_float(data.get("background_match_ratio")),
            background_complexity=_opt_float(data.get("background_complexity")),
            color_confidence=float(data.get("color_confidence", 0.0)),
            method=data.get("method", "unsupported"),
            unsupported_reason=data.get("unsupported_reason"),
        )


@dataclass
class ArtWarning:
    code: str
    message: str

    def to_dict(self) -> dict[str, str]:
        return {"code": self.code, "message": self.message}


@dataclass
class ArtFinding:
    rule_id: str
    dimension: str
    severity: Severity
    message: str
    confidence: float
    slide_index: int | None = None
    primary: ArtElementRef | None = None
    related: list[ArtElementRef] = field(default_factory=list)
    details: dict[str, JsonValue] = field(default_factory=dict)
    evidence_sources: frozenset[str] = frozenset()
    evidence_reliability: float | None = None
    evidence_method: str | None = None
    severity_override: bool = False
    severity_override_source: Literal["user", "feedback"] | None = None
    experimental: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "rule_id": self.rule_id,
            "dimension": self.dimension,
            "severity": self.severity.name,
            "message": self.message,
            "confidence": self.confidence,
            "slide_index": self.slide_index,
            "primary": self.primary.to_dict() if self.primary else None,
            "related": [r.to_dict() for r in self.related],
            "details": self.details,
            "evidence_sources": sorted(self.evidence_sources),
            "evidence_reliability": self.evidence_reliability,
            "evidence_method": self.evidence_method,
            **({"severity_override": True} if self.severity_override else {}),
            **(
                {"severity_override_source": self.severity_override_source}
                if self.severity_override_source is not None
                else {}
            ),
            **({"experimental": True} if self.experimental else {}),
        }


@dataclass
class DimensionAssessment:
    dimension: str
    status: AssessmentStatus = "assessed"
    grade: Grade | None = None
    confidence: float = 0.0
    evidence_coverage: float = 1.0
    findings: list[ArtFinding] = field(default_factory=list)
    warnings: list[ArtWarning] = field(default_factory=list)
    reliability: float | None = None
    minimum_reliability: float | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "dimension": self.dimension,
            "status": self.status,
            "grade": self.grade,
            "confidence": self.confidence,
            "evidence_coverage": self.evidence_coverage,
            "findings": [f.to_dict() for f in self.findings],
            "warnings": [w.to_dict() for w in self.warnings],
            **({"reliability": self.reliability} if self.reliability is not None else {}),
            **(
                {"minimum_reliability": self.minimum_reliability}
                if self.minimum_reliability is not None
                else {}
            ),
        }


@dataclass
class ArtSlideReport:
    slide_index: int
    dimensions: list[DimensionAssessment] = field(default_factory=list)
    dominant_focus: dict[str, Any] | None = None
    visual_balance: dict[str, Any] | None = None

    def by_dimension(self, dim: str) -> DimensionAssessment | None:
        return next((d for d in self.dimensions if d.dimension == dim), None)

    def to_dict(self) -> dict[str, Any]:
        return {
            "slide_index": self.slide_index,
            "dimensions": [d.to_dict() for d in self.dimensions],
            "dominant_focus": self.dominant_focus,
            "visual_balance": self.visual_balance,
        }


@dataclass
class ArtReport:
    schema_version: str = ART_REPORT_SCHEMA_VERSION
    profile: str = "balanced"
    slides: list[ArtSlideReport] = field(default_factory=list)
    deck_findings: list[ArtFinding] = field(default_factory=list)
    experimental_score: float | None = None
    experimental_score_mode: Literal["grade_mean", "worth_sigmoid"] | None = None  # #130：分数来源
    warnings: list[ArtWarning] = field(default_factory=list)  # #151：feedback 级模型告警

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "profile": self.profile,
            "slides": [s.to_dict() for s in self.slides],
            "deck_findings": [f.to_dict() for f in self.deck_findings],
            "experimental_score": self.experimental_score,
            "experimental_score_mode": self.experimental_score_mode,
            "warnings": [w.to_dict() for w in self.warnings],
        }


@dataclass
class DeckQualityReport:
    geometry: object | None = None
    art: ArtReport | None = None
    warnings: list[ArtWarning] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "geometry": (
                self.geometry.to_dict() if hasattr(self.geometry, "to_dict") else self.geometry
            ),
            "art": self.art.to_dict() if self.art else None,
            "warnings": [w.to_dict() for w in self.warnings],
        }
