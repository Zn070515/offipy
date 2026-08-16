"""层级规则：标题过小 / 无视觉焦点。"""

from __future__ import annotations

from typing import TYPE_CHECKING

from offipy.audit import Severity

from .profiles import RULE_NO_FOCUS, RULE_TITLE_TOO_SMALL
from .rules import RuleContext, RuleEvaluation, RuleSpec, make_finding

if TYPE_CHECKING:
    from .models import ArtElement, ArtSlide


def _title_el(slide: ArtSlide) -> ArtElement | None:
    return next((e for e in slide.elements if e.role == "title"), None)


def title_too_small_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    t = _title_el(slide)
    eligible = 1 if t is not None else 0
    if t is None:
        return RuleEvaluation(covered_count=0, eligible_count=0)
    n = t.font_size_norm
    if n is None:
        return RuleEvaluation(covered_count=0, eligible_count=eligible)
    if n >= ctx.profile.title_size_min_norm:
        return RuleEvaluation(covered_count=1, eligible_count=eligible)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_TITLE_TOO_SMALL,
                "hierarchy",
                Severity.MID,
                f"标题字号过小（font_size_norm={n:.4f}）。",
                0.55,
                slide.index,
                primary=t,
                details={
                    "font_size_norm": round(n, 4),
                    "ratio_vs_min": round(n / ctx.profile.title_size_min_norm, 3),
                },
            )
        ],
        covered_count=1,
        eligible_count=eligible,
    )


def no_focus_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    focus = ctx.features.get("focus", {})
    els = [e for e in slide.elements if e.role not in ("background", "container", "decoration")]
    if len(els) < 3:
        # 元素不足 → 规则不适用，不进 coverage（不拖累整维证据）
        return RuleEvaluation(covered_count=0, eligible_count=0)
    eligible = len(els)
    covered = len([e for e in els if e.font_size_norm is not None])
    if focus.get("has_focus"):
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    dominant = max(els, key=lambda e: e.area)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_NO_FOCUS,
                "hierarchy",
                Severity.LOW,
                "页面没有明确的视觉焦点（字号层级过平）。",
                0.25,
                slide.index,
                primary=dominant,
                details={"focus_ratio": focus.get("ratio")},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


RULES = [
    RuleSpec(rule_id=RULE_TITLE_TOO_SMALL, dimension="hierarchy", run=title_too_small_rule),
    RuleSpec(rule_id=RULE_NO_FOCUS, dimension="hierarchy", run=no_focus_rule, experimental=True),
]
