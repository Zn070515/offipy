"""分析编排：analyze_scene（单场景）与 analyze_deck（几何 + 艺术组合）。

rev2.1：
- coverage 由每条规则的 RuleEvaluation 汇总，不猜常数；
- analyze_deck 先 audit_pptx 一次，build_scene(measurements=, pptx_report=geometry) 复用；
- PPTX-only 可产出艺术报告（typography/color 多为 insufficient_evidence）。
"""

from __future__ import annotations

from offipy.audit import audit_pptx
from offipy.exceptions import InvalidArgumentError

from .adapters import build_scene
from .color import RULES as COLOR_RULES
from .composition import RULES as COMPOSITION_RULES
from .consistency import assess_deck
from .features import compute_features
from .hierarchy import RULES as HIERARCHY_RULES
from .media import RULES as MEDIA_RULES
from .models import (
    ART_REPORT_SCHEMA_VERSION,
    ArtReport,
    ArtScene,
    ArtSlideReport,
    ArtWarning,
    DeckQualityReport,
    DimensionAssessment,
)
from .profiles import ArtProfile, get_profile
from .rules import RuleContext, assess_dimension
from .typography import RULES as TYPOGRAPHY_RULES

_DIMENSION_RULES = {
    "hierarchy": HIERARCHY_RULES,
    "composition": COMPOSITION_RULES,
    "typography": TYPOGRAPHY_RULES,
    "color": COLOR_RULES,
    "media": MEDIA_RULES,
}

_DIM_ORDER = ("hierarchy", "composition", "typography", "color", "media")

_GRADE_SCORE = {"excellent": 0.0, "good": 0.3, "attention": 0.6, "poor": 1.0}


def _experimental_score(report: ArtReport) -> float | None:
    """规则评级均值（0-100），仅当 ≥3 个 assessed 维度才返回。"""
    assessed = 0
    total = 0.0
    for s in report.slides:
        for d in s.dimensions:
            if d.status != "assessed" or d.grade is None:
                continue
            total += _GRADE_SCORE[d.grade]
            assessed += 1
    if assessed < 3:
        return None
    score = (1.0 - total / assessed) * 100.0
    return round(score, 1)


def analyze_scene(
    scene: ArtScene,
    profile: str | ArtProfile | None = None,
    include_experimental_score: bool = False,
) -> ArtReport:
    prof = profile if isinstance(profile, ArtProfile) else get_profile(profile or "balanced")
    report = ArtReport(schema_version=ART_REPORT_SCHEMA_VERSION, profile=prof.name)
    for slide in scene.slides:
        feats = compute_features(slide)
        ctx = RuleContext(
            profile=prof,
            slide=slide,
            slide_index=slide.index,
            features=feats,
            deck=scene,
            sources=frozenset(scene.sources),
        )
        dims: list[DimensionAssessment] = []
        for dim in _DIM_ORDER:
            dims.append(assess_dimension(dim, _DIMENSION_RULES[dim], ctx))
        report.slides.append(
            ArtSlideReport(
                slide_index=slide.index,
                dimensions=dims,
                dominant_focus=feats.get("focus"),
                visual_balance=feats.get("mass"),
            )
        )
    report.deck_findings = assess_deck(scene, prof)
    if include_experimental_score:
        report.experimental_score = _experimental_score(report)
    return report


def analyze_deck(
    *,
    pptx: str | None = None,
    measurements: str | dict | None = None,
    slides_dir: str | None = None,
    profile: str | ArtProfile | None = None,
    include_experimental_score: bool = False,
) -> DeckQualityReport:
    """组合入口：几何审计（可选）+ 像素（可选）+ 艺术分析（可选）。"""
    if measurements is None and pptx is None and slides_dir is None:
        raise InvalidArgumentError("analyze_deck requires measurements, pptx, or slides_dir")
    geometry = None
    if pptx:
        geometry = audit_pptx(pptx)
    art = None
    warnings: list[ArtWarning] = []
    if measurements is not None or pptx is not None or slides_dir is not None:
        scene = build_scene(measurements=measurements, pptx_report=geometry, slides_dir=slides_dir)
        art = analyze_scene(
            scene, profile=profile, include_experimental_score=include_experimental_score
        )
        warnings = list(scene.warnings)
        if pptx is not None and measurements is None and slides_dir is None:
            warnings.append(
                ArtWarning(
                    code="art.evidence.limited",
                    message="仅 pptx 源：无像素/字号证据，依赖 font_size 的维度可能证据不足",
                )
            )
    return DeckQualityReport(geometry=geometry, art=art, warnings=warnings)
