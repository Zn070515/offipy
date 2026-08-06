"""art 数据模型（v0.12 契约冻结，rev2.1 一次定死全部字段）。

公开 slide 索引一律 1-based；坐标/尺寸按页宽高归一化但不限 [0,1]；
字号规范：font_size + font_size_unit + font_size_norm（普通字段，适配器直接填），
规则只消费 font_size_norm，模型不做任何 norm 回填。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from offipy.audit import Severity

ART_SCHEMA_VERSION = "0.1"
ART_REPORT_SCHEMA_VERSION = "0.1"

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
    def from_dict(cls, data: dict) -> ArtColor:
        return cls(int(data["r"]), int(data["g"]), int(data["b"]), float(data.get("a", 1.0)))

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

    def to_dict(self) -> dict:
        return {"r": self.r, "g": self.g, "b": self.b, "a": self.a}


@dataclass(frozen=True)
class ArtTextRun:
    text: str
    font_size: float | None = None
    font_size_unit: str = "unknown"
    font_family: str | None = None
    color: ArtColor | None = None  # 该 run 的前景色（文本前景）

    def to_dict(self) -> dict:
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

    def to_dict(self) -> dict:
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
    evidence: dict = field(default_factory=dict)
    container: bool = False
    decoration: bool = False

    @property
    def area(self) -> float:
        return self.width * self.height

    def has_text(self) -> bool:
        return bool(self.text and self.text.strip())

    def to_dict(self) -> dict:
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
        }


@dataclass
class ArtSlide:
    index: int  # 1-based
    width: float
    height: float
    elements: list[ArtElement] = field(default_factory=list)
    background_color: ArtColor | None = None

    def by_id(self, element_id: str) -> ArtElement | None:
        return next((e for e in self.elements if e.element_id == element_id), None)


@dataclass
class ArtScene:
    slides: list[ArtSlide] = field(default_factory=list)
    width_unit: str = "px"  # "px" | "pt"
    warnings: list[ArtWarning] = field(default_factory=list)
    sources: set[str] = field(default_factory=set)

    @classmethod
    def from_dict(cls, data: dict) -> ArtScene:
        """手写/测试场景输入协议（与 MeasurementAdapter 的真实格式分离）。"""
        slides: list[ArtSlide] = []
        for i, s in enumerate(data.get("slides", [])):
            idx = int(s.get("index") or i + 1)
            height = float(s.get("height", 1080.0))
            elements = [_element_from_dict(e, idx, height) for e in s.get("elements", [])]
            bg = s.get("background_color")
            slides.append(
                ArtSlide(
                    index=idx,
                    width=float(s.get("width", 1920.0)),
                    height=height,
                    elements=elements,
                    background_color=ArtColor.from_dict(bg) if bg else None,
                )
            )
        return cls(
            slides=slides,
            width_unit=data.get("width_unit", "px"),
            warnings=[ArtWarning(w["code"], w["message"]) for w in data.get("warnings", [])],
            sources=set(data.get("sources", [])),
        )

    def by_slide(self, index: int) -> ArtSlide | None:
        return next((s for s in self.slides if s.index == index), None)


def _element_from_dict(data: dict, slide_index: int, height: float) -> ArtElement:
    fg = data.get("foreground") or data.get("color")
    bg = data.get("background")
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
            container=bool(data.get("container", False)),
            decoration=bool(data.get("decoration", False)),
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
        is_background=bool(data.get("is_background", False)),
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
        container=bool(data.get("container", False)),
        decoration=bool(data.get("decoration", False)),
    )


def _norm_of(data: dict, height: float) -> float | None:
    """手写场景的 norm：显式提供用显式值；否则 px/pt 同单位直接除。"""
    explicit = data.get("font_size_norm")
    if explicit is not None:
        return float(explicit)
    fs = data.get("font_size")
    unit = data.get("font_size_unit", "unknown")
    if fs is None or unit == "unknown" or not height:
        return None
    return float(fs) / height


def _color_from(c) -> ArtColor | None:
    if c is None:
        return None
    if isinstance(c, ArtColor):
        return c
    if isinstance(c, dict):
        return ArtColor.from_dict(c)
    if isinstance(c, (list, tuple)) and len(c) >= 3:
        return ArtColor(int(c[0]), int(c[1]), int(c[2]), float(c[3]) if len(c) > 3 else 1.0)
    return None


@dataclass
class ArtWarning:
    code: str
    message: str

    def to_dict(self) -> dict:
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

    def to_dict(self) -> dict:
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

    def to_dict(self) -> dict:
        return {
            "dimension": self.dimension,
            "status": self.status,
            "grade": self.grade,
            "confidence": self.confidence,
            "evidence_coverage": self.evidence_coverage,
            "findings": [f.to_dict() for f in self.findings],
            "warnings": [w.to_dict() for w in self.warnings],
        }


@dataclass
class ArtSlideReport:
    slide_index: int
    dimensions: list[DimensionAssessment] = field(default_factory=list)
    dominant_focus: dict | None = None
    visual_balance: dict | None = None

    def by_dimension(self, dim: str) -> DimensionAssessment | None:
        return next((d for d in self.dimensions if d.dimension == dim), None)

    def to_dict(self) -> dict:
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

    def to_dict(self) -> dict:
        return {
            "schema_version": self.schema_version,
            "profile": self.profile,
            "slides": [s.to_dict() for s in self.slides],
            "deck_findings": [f.to_dict() for f in self.deck_findings],
            "experimental_score": self.experimental_score,
        }


@dataclass
class DeckQualityReport:
    geometry: object | None = None
    art: ArtReport | None = None
    warnings: list[ArtWarning] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "geometry": (
                self.geometry.to_dict() if hasattr(self.geometry, "to_dict") else self.geometry
            ),
            "art": self.art.to_dict() if self.art else None,
            "warnings": [w.to_dict() for w in self.warnings],
        }
