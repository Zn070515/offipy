"""构图规则：失衡 / 角落聚集 / 间距漂移。"""

from __future__ import annotations

from offipy.audit import Severity

from .models import ArtElement, ArtSlide, ArtWarning
from .profiles import (
    RULE_BACKGROUND_LIKE_AREA,
    RULE_CORNER_CLUSTER,
    RULE_OFF_BALANCE,
    RULE_SPACING_DRIFT,
)
from .rules import RuleContext, RuleEvaluation, RuleSpec, make_finding

_SKIP_ROLES = {"background", "container", "decoration", "page_number", "footer"}

_BG_CONF_MIN = 0.7
_BG_UNIFORMITY_MIN = 0.7
_MAX_OCCUPANCY = 0.5
_FULL_BLEED_IMAGE_AREA = 0.9


def _weighted(elements: list[ArtElement]) -> list[ArtElement]:
    return [e for e in elements if e.role not in _SKIP_ROLES and not e.is_background]


def _quadrant_mass(elements: list[ArtElement]) -> dict[str, float]:
    """四象限质量占比（TL/TR/BL/BR，归一化坐标）。"""
    q = {"TL": 0.0, "TR": 0.0, "BL": 0.0, "BR": 0.0}
    total = 0.0
    for e in elements:
        cx, cy = e.x + e.width / 2, e.y + e.height / 2
        if cx < 0.5 and cy < 0.5:
            q["TL"] += e.area
        elif cx >= 0.5 and cy < 0.5:
            q["TR"] += e.area
        elif cx < 0.5 and cy >= 0.5:
            q["BL"] += e.area
        else:
            q["BR"] += e.area
        total += e.area
    if total:
        q = {k: v / total for k, v in q.items()}
    return q


def off_balance_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    els = _weighted(slide.elements)
    eligible = covered = len(els)
    if len(els) < 3:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    mass = ctx.features.get("mass", {})
    total = sum(e.area for e in els) or 1.0
    cx = sum((e.x + e.width / 2) * e.area for e in els) / total
    dist = abs(cx - 0.5)
    if dist <= ctx.profile.balance_tol:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    dominant = max(els, key=lambda e: e.area)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_OFF_BALANCE,
                "composition",
                Severity.MID,
                f"视觉重心偏一侧（偏离 0.5 达 {dist:.3f}）。",
                0.4,
                slide.index,
                primary=dominant,
                details={"balance_dist": round(dist, 3), "ink": mass.get("ink")},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


def corner_cluster_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    els = _weighted(slide.elements)
    eligible = covered = len(els)
    if len(els) < 3:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    q = _quadrant_mass(els)
    mx = max(q.values())
    if mx <= ctx.profile.corner_cluster_ratio:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    dominant = max(els, key=lambda e: e.area)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_CORNER_CLUSTER,
                "composition",
                Severity.LOW,
                f"内容集中在单个角落（{max(q, key=q.__getitem__)} 占 {mx:.2f}）。",
                0.4,
                slide.index,
                primary=dominant,
                details={"quadrants": q},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


def spacing_drift_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    sp = ctx.features.get("spacing", {})
    h, v = sp.get("horizontal", {}), sp.get("vertical", {})
    h_ok = (
        h.get("drift_count", 0) >= 1
        and h.get("max_drift_ratio", 0.0) >= ctx.profile.spacing_drift_tol
    )
    v_ok = (
        v.get("drift_count", 0) >= 1
        and v.get("max_drift_ratio", 0.0) >= ctx.profile.spacing_drift_tol
    )
    eligible = covered = 1 if sp else 0
    if not (h_ok or v_ok):
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_SPACING_DRIFT,
                "composition",
                Severity.MID,
                "元素间距不匀，存在明显漂移。",
                0.45,
                slide.index,
                details={"horizontal": h, "vertical": v},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


def _full_bleed_image(slide: ArtSlide) -> bool:
    return any(e.kind == "image" and e.area >= _FULL_BLEED_IMAGE_AREA for e in slide.elements)


def background_like_area_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    """页面级留白提示（experimental）：像素背景证据 + 联合条件。

    联合条件：背景置信 ≥0.7 ∧ 均匀度 ≥0.7 ∧ 元素占用 ≤0.5 ∧ 非全幅图片。
    低置信但高相似面积 → warning 不提示；占用/全图 → 静默跳过。
    """
    pe = slide.pixel_evidence
    eligible = 1 if pe is not None else 0
    if pe is None or pe.background_like_ratio is None:
        return RuleEvaluation(covered_count=0, eligible_count=eligible)
    if pe.background_confidence is None or pe.background_uniformity is None:
        return RuleEvaluation(covered_count=0, eligible_count=eligible)
    conf_ok = (
        pe.background_confidence >= _BG_CONF_MIN and pe.background_uniformity >= _BG_UNIFORMITY_MIN
    )
    if not conf_ok:
        if pe.background_like_ratio > ctx.profile.max_background_like_ratio:
            return RuleEvaluation(
                covered_count=0,
                eligible_count=eligible,
                warnings=[
                    ArtWarning(
                        code="art.pixel.background_low_confidence",
                        message="背景相似面积高但背景像素证据置信不足，不提示留白",
                    )
                ],
            )
        return RuleEvaluation(covered_count=0, eligible_count=eligible)
    occupancy = (ctx.features.get("density") or {}).get("union_area_ratio", 1.0)
    if occupancy > _MAX_OCCUPANCY or _full_bleed_image(slide):
        return RuleEvaluation(covered_count=0, eligible_count=eligible)
    if pe.background_like_ratio <= ctx.profile.max_background_like_ratio:
        return RuleEvaluation(covered_count=1, eligible_count=eligible)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_BACKGROUND_LIKE_AREA,
                "composition",
                Severity.LOW,
                f"页面大面积近似背景色（占比 {pe.background_like_ratio:.2f}），可能留白过多。",
                0.4,
                slide.index,
                evidence_sources={"pixel"},
                evidence_reliability=0.5,
                evidence_method=None,
                details={"background_like_ratio": round(pe.background_like_ratio, 3)},
            )
        ],
        covered_count=1,
        eligible_count=eligible,
        reliability=0.5,
    )


RULES = [
    RuleSpec(
        rule_id=RULE_OFF_BALANCE, dimension="composition", run=off_balance_rule, experimental=True
    ),
    RuleSpec(
        rule_id=RULE_CORNER_CLUSTER,
        dimension="composition",
        run=corner_cluster_rule,
        experimental=True,
    ),
    RuleSpec(rule_id=RULE_SPACING_DRIFT, dimension="composition", run=spacing_drift_rule),
    RuleSpec(
        rule_id=RULE_BACKGROUND_LIKE_AREA,
        dimension="composition",
        run=background_like_area_rule,
        experimental=True,
    ),
]
