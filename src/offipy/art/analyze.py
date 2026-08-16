"""分析编排：analyze_scene（单场景）与 analyze_deck（几何 + 艺术组合）。

rev2.1：
- coverage 由每条规则的 RuleEvaluation 汇总，不猜常数；
- analyze_deck 先 audit_pptx 一次，build_scene(measurements=, pptx_report=geometry) 复用；
- PPTX-only 可产出艺术报告（typography/color 多为 insufficient_evidence）。
"""

from __future__ import annotations

import dataclasses
from collections import Counter
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
from .rules import RuleContext, apply_profile_to_finding, assess_dimension, grade_from_findings
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

# #111：规则需 ≥3 条有效标签才允许被 severity_shift（模型仍全特征预测，
# 但 shift 的「应用」按规则证据门禁——0 标签规则不做跨规则泛化 shift）
_MIN_LABELS_PER_RULE = 3

# #158：quality_score 需要 ≥3 个 assessed 维度 worth——单点/证据不足不冒充分数
_MIN_QUALITY_SCORE_SAMPLES = 3


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


def _reconcile_grades(report: ArtReport) -> None:
    """#132：severity_shift 改了 finding.severity → 重推 assessed 维度 grade。

    grade 的单一事实来源是 rules.grade_from_findings；学习 pass 在 shift 后调用，
    保证 grade 与 post-shift finding 一致（confidence / evidence_coverage 不动）。
    """
    for slide in report.slides:
        for dim in slide.dimensions:
            if dim.status == "assessed" and dim.findings:
                dim.grade = grade_from_findings(dim.findings)


def _apply_learning_pass(
    report: ArtReport,
    scene: ArtScene,
    profile_name: str,
    feedback_dir: str | Path | None,
    *,
    want_score: bool,
) -> None:
    """学习后处理 pass：severity_shift（severity_override=False 才作用）+ quality.score。

    severity_shift 是「推荐」语义：先改 finding.severity，再按 post-shift findings
    重推 assessed 维度 grade（#132 _reconcile_grades，避免 grade 与 finding 不一致）。
    quality.score 有请求且模型有效时替换 experimental_score。全程惰性 import：
    无 feedback extra 时 ImportError → 跳过。
    """
    try:
        from offipy.art.features_registry import encode_features
        from offipy.art.feedback import load_records
        from offipy.feedback import infer
        from offipy.feedback.heads import apply_severity_shift, severity_shift_from_worth
        from offipy.feedback.pairs import valid_records
    except ImportError:
        return
    bundle = infer.ModelBundle.load(feedback_dir)
    if bundle is None:
        # #158：显式反馈路径无有效模型 → 不再静默回退 v2，发 warning 让用户知情。
        report.warnings.append(
            ArtWarning(
                code="feedback.model.unavailable",
                message="指定反馈目录无有效学习模型（缺失/过期/损坏），本次分析回退 v2 规则",
            )
        )
        return
    if bundle.saturation:
        report.warnings.append(
            ArtWarning(
                code="feedback.model.saturated",
                message="反馈模型输出饱和（样本间 quality_score 跨度过小，判别力不足），"
                "建议补充更多差异样本后重训",
            )
        )
    # #111：按规则证据门禁——只有有效标签 ≥ _MIN_LABELS_PER_RULE 的规则才允许
    # severity_shift。模型仍全特征预测（跨规则泛化保留在 worth 计算里），
    # 但 shift 的「应用」被门禁挡住，0 标签规则不做跨规则泛化 shift。
    labeled = Counter(
        r.rule_id for r in valid_records(load_records(feedback_dir), profile=profile_name)
    )
    # #133：quality_score 人口统计——covered 是进入分数计算的 assessed 维度 worth 子集；
    # abstain/ood 被挡下的 finding 单独计数，让「分数只基于置信子集」的偏差可见。
    covered: list[float] = []
    total_findings = 0
    abstained_count = 0
    ood_count = 0
    for finding, slide_index in _all_findings(report):
        if finding.severity_override:
            continue  # user / feedback override 的 finding 一律跳过
        if labeled[finding.rule_id] < _MIN_LABELS_PER_RULE:
            continue  # 证据不足的规则不 shift（Counter 对未出现 key 返回 0）
        total_findings += 1
        slide = scene.by_slide(slide_index) if slide_index is not None else None
        feats = encode_features(finding, slide, scene, profile_name)
        if bundle.should_abstain(feats):
            abstained_count += 1
            continue  # 保守：模型不确定的 finding 不 shift（回退 v2）
        if bundle.ood_flagged(feats):
            ood_count += 1
            continue  # 特征 OOD 的 finding 不 shift（回退 v2）
        worth = infer.model_worth(feats, bundle)  # 模块级 seam，测试 monkeypatch 生效
        shift = severity_shift_from_worth(worth)
        before = finding.severity
        finding.severity = apply_severity_shift(before, shift)
        if finding.severity != before:
            # #157：severity_shift 必须有 provenance——标 override + 来源 + worth/delta/head
            finding.severity_override = True
            finding.severity_override_source = "feedback"
            finding.details["feedback"] = {
                "head": "severity_shift",
                "worth": round(worth, 4),
                "shift": round(shift, 4),
                "before": before.name,
                "after": finding.severity.name,
            }
        # #158 assessed 门：维度状态在 report 侧（ArtSlideReport.by_dimension）——scene 的
        # ArtSlide 无 by_dimension；仅 assessed 维度的 slide finding 才计入 quality_score 人口。
        if slide_index is not None:
            slide_report = next((sr for sr in report.slides if sr.slide_index == slide_index), None)
            if slide_report is not None:
                dim = slide_report.by_dimension(finding.dimension)
                if dim is not None and dim.status == "assessed":
                    covered.append(worth)
    if want_score and len(covered) >= _MIN_QUALITY_SCORE_SAMPLES:
        mean_worth = sum(covered) / len(covered)
        confident_quality = bundle.quality_score(mean_worth)
        report.experimental_score = confident_quality
        report.experimental_score_mode = "worth_sigmoid"  # #130：区分 grade-mean 来源
        report.quality_score_coverage = {
            "mean_worth": round(mean_worth, 4),
            "confident_quality": confident_quality,
            "covered_findings": len(covered),
            "total_findings": total_findings,
            "abstained_count": abstained_count,
            "ood_count": ood_count,
        }
    # #132：severity_shift 改 severity 后重推 grade，否则 grade 与 finding 不一致
    _reconcile_grades(report)


def analyze_scene(
    scene: ArtScene,
    profile: str | ArtProfile | None = None,
    include_experimental_score: bool = False,
    *,
    feedback: bool = False,
    feedback_dir: str | Path | None = None,
) -> ArtReport:
    if feedback and not feedback_dir:
        raise InvalidArgumentError(
            "feedback=True 必须显式提供 feedback_dir（学习路径只读指定目录，"
            "不静默加载全局 ~/.offipy 模型）"
        )
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
    # #159：rule.delta 落到报告（profile 已被 _resolve_profile 替换为学习调整后的副本）
    report.feedback_adjustments = dict(prof.feedback_severity_adjustments)
    if include_experimental_score:
        report.experimental_score = _experimental_score(report)
        if report.experimental_score is not None:
            report.experimental_score_mode = "grade_mean"  # #130：分数来源标注
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
        warnings = list(scene.warnings) + list(art.warnings if art else [])
        if pptx is not None and measurements is None and slides_dir is None:
            warnings.append(
                ArtWarning(
                    code="art.evidence.limited",
                    message="仅 pptx 源：无像素/字号证据，依赖 font_size 的维度可能证据不足",
                )
            )
    return DeckQualityReport(geometry=geometry, art=art, warnings=warnings)
