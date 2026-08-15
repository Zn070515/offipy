"""分析编排：analyze_scene（单场景）与 analyze_deck（几何 + 艺术组合）。

rev2.1：
- coverage 由每条规则的 RuleEvaluation 汇总，不猜常数；
- analyze_deck 先 audit_pptx 一次，build_scene(measurements=, pptx_report=geometry) 复用；
- PPTX-only 可产出艺术报告（typography/color 多为 insufficient_evidence）。
"""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING, Any

from offipy.audit import audit_pptx
from offipy.exceptions import InvalidArgumentError

from .adapters import build_scene
from .color import RULES as COLOR_RULES
from .composition import RULES as COMPOSITION_RULES
from .consistency import assess_deck
from .features import compute_features
from .feedback import apply_feedback
from .hierarchy import RULES as HIERARCHY_RULES
from .media import RULES as MEDIA_RULES
from .models import (
    ART_REPORT_SCHEMA_VERSION,
    ArtFinding,
    ArtReport,
    ArtScene,
    ArtSlideReport,
    ArtWarning,
    DeckQualityReport,
    DimensionAssessment,
)
from .profiles import ArtProfile, get_profile
from .rules import RuleContext, apply_profile_to_finding, assess_dimension
from .typography import RULES as TYPOGRAPHY_RULES

if TYPE_CHECKING:
    from collections.abc import Iterator
    from pathlib import Path

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


def _learned_adjustments_safe(
    profile_name: str, feedback_dir: str | Path | None
) -> dict[str, int] | None:
    """有效模型 → learned rule.delta；无模型/无 numpy/任何异常 → None（回退 v2）。"""
    try:
        from offipy.feedback.infer import learned_adjustments
    except ImportError:
        return None
    try:
        return learned_adjustments(profile_name, feedback_dir=feedback_dir)
    except Exception:
        return None


def _resolve_profile(
    profile: str | ArtProfile | None,
    *,
    feedback: bool,
    feedback_dir: str | Path | None,
) -> ArtProfile:
    prof = profile if isinstance(profile, ArtProfile) else get_profile(profile or "balanced")
    if not feedback:
        return prof
    adjustments = _learned_adjustments_safe(prof.name, feedback_dir)
    if adjustments is not None:
        # F2-B：有效模型下学习路径是 feedback_severity_adjustments 的唯一来源，
        # 不再叠加 v2 手写调整
        return dataclasses.replace(prof, feedback_severity_adjustments=adjustments)
    return apply_feedback(prof, feedback_dir=feedback_dir)


def _all_findings(report: ArtReport) -> Iterator[tuple[ArtFinding, int | None]]:
    """展平 slide + deck 的全部 finding（带 slide 上下文）。"""
    for slide_report in report.slides:
        for dim in slide_report.dimensions:
            for f in dim.findings:
                yield f, slide_report.slide_index
    for f in report.deck_findings:
        yield f, None


def _apply_learning_pass(
    report: ArtReport,
    scene: ArtScene,
    profile_name: str,
    feedback_dir: str | Path | None,
    *,
    want_score: bool,
) -> None:
    """学习后处理 pass：severity_shift（severity_override=False 才作用）+ quality.score。

    severity_shift 是「推荐」语义：只改 finding.severity，不改 dimension grade
    （grade 在 assess_dimension 时已定）。quality.score 有请求且模型有效时替换
    experimental_score。全程惰性 import：无 feedback extra 时 ImportError → 跳过。
    """
    try:
        from offipy.art.features_registry import (
            encode_features,
            feature_keys,
            feature_schema_version,
        )
        from offipy.feedback import infer
        from offipy.feedback.heads import apply_severity_shift, severity_shift_from_worth
        from offipy.feedback.model import load_model, model_file, model_valid, weights_from_dict
    except ImportError:
        return
    data = load_model(model_file(feedback_dir))
    if data is None or not model_valid(data, feature_schema_version()):
        return
    try:
        mlp = weights_from_dict(
            data, input_dim=len(feature_keys()), hidden_dims=tuple(data["hidden_dims"])
        )
    except (ValueError, KeyError, TypeError):
        return  # 损坏模型（schema 匹配但权重形状错）→ 视为无模型，回退 v2
    worths: list[float] = []
    for finding, slide_index in _all_findings(report):
        if finding.severity_override:
            continue  # user / feedback override 的 finding 一律跳过
        slide = scene.by_slide(slide_index) if slide_index is not None else None
        feats = encode_features(finding, slide, scene, profile_name)
        worth = infer.model_worth(feats, mlp)
        shift = severity_shift_from_worth(worth)
        finding.severity = apply_severity_shift(finding.severity, shift)
        worths.append(worth)
    if want_score and worths:
        mean_worth = sum(worths) / len(worths)
        report.experimental_score = infer.quality_score_for_report(mean_worth)


def analyze_scene(
    scene: ArtScene,
    profile: str | ArtProfile | None = None,
    include_experimental_score: bool = False,
    *,
    feedback: bool = False,
    feedback_dir: str | Path | None = None,
) -> ArtReport:
    prof = _resolve_profile(profile, feedback=feedback, feedback_dir=feedback_dir)
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
        dims: list[DimensionAssessment] = [
            assess_dimension(dim, _DIMENSION_RULES[dim], ctx) for dim in _DIM_ORDER
        ]
        report.slides.append(
            ArtSlideReport(
                slide_index=slide.index,
                dimensions=dims,
                dominant_focus=feats.get("focus"),
                visual_balance=feats.get("mass"),
            )
        )
    report.deck_findings = [apply_profile_to_finding(f, prof) for f in assess_deck(scene, prof)]
    if include_experimental_score:
        report.experimental_score = _experimental_score(report)
    if feedback:
        _apply_learning_pass(
            report, scene, prof.name, feedback_dir, want_score=include_experimental_score
        )
    return report


def analyze_deck(
    *,
    pptx: str | None = None,
    measurements: str | dict[str, Any] | None = None,
    slides_dir: str | None = None,
    profile: str | ArtProfile | None = None,
    include_experimental_score: bool = False,
    feedback: bool = False,
    feedback_dir: str | Path | None = None,
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
            scene,
            profile=profile,
            include_experimental_score=include_experimental_score,
            feedback=feedback,
            feedback_dir=feedback_dir,
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
