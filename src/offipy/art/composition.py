"""构图规则：失衡 / 角落聚集 / 间距漂移。"""

from __future__ import annotations

from offipy.audit import Severity

from .models import ArtElement, ArtSlide
from .profiles import (
    RULE_CORNER_CLUSTER,
    RULE_OFF_BALANCE,
    RULE_SPACING_DRIFT,
)
from .rules import RuleContext, RuleEvaluation, RuleSpec, make_finding

_SKIP_ROLES = {"background", "container", "decoration", "page_number", "footer"}


def _weighted(elements: list[ArtElement]) -> list[ArtElement]:
    return [e for e in elements if e.role not in _SKIP_ROLES and not e.is_background]


def _quadrant_mass(elements: list[ArtElement]) -> dict:
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
                0.25,
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
                0.3,
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
]
