"""字体排版规则：多字体 / 过小文字 / 无层级。

rev2.1：返回 RuleEvaluation；covered/eligible 表达「有多少对象有证据被评估」。
"""

from __future__ import annotations

from collections import Counter
from typing import TYPE_CHECKING

from offipy.audit import Severity

from .features import font_hierarchy
from .profiles import (
    RULE_FLAT_SCALE,
    RULE_MANY_FAMILIES,
    RULE_TINY_TEXT,
)
from .rules import (
    RuleContext,
    RuleEvaluation,
    RuleSpec,
    make_finding,
)

if TYPE_CHECKING:
    from .models import ArtElement, ArtSlide


def _rarest_family_el(slide: ArtSlide) -> ArtElement | None:
    counts: Counter[str] = Counter()
    for e in slide.elements:
        for r in e.runs:
            if r.font_family:
                counts[r.font_family] += 1
    if not counts:
        return None
    rarest = min(counts, key=lambda f: counts[f])
    return next((e for e in slide.elements if any(r.font_family == rarest for r in e.runs)), None)


def many_families_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    texts = [e for e in slide.elements if e.has_text()]
    eligible = texts  # 所有文本元素都在评估范围
    covered = [e for e in texts if any(r.font_family for r in e.runs)]
    families = set()
    for e in texts:
        for r in e.runs:
            if r.font_family:
                families.add(r.font_family)
    if len(families) <= ctx.profile.max_font_families:
        return RuleEvaluation(covered_count=len(covered), eligible_count=len(eligible))
    if len(covered) / len(eligible) < 0.5:
        return RuleEvaluation(covered_count=len(covered), eligible_count=len(eligible))
    primary = _rarest_family_el(slide)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_MANY_FAMILIES,
                "typography",
                Severity.MID,
                f"页面使用 {len(families)} 种字体，超过 {ctx.profile.max_font_families} 的上限。",
                0.6,
                slide.index,
                primary=primary,
                details={"families": sorted(families)},
            )
        ],
        covered_count=len(covered),
        eligible_count=len(eligible),
    )


def tiny_text_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    out = []
    eligible = [e for e in slide.elements if e.has_text()]
    covered = 0
    for e in eligible:
        n = e.font_size_norm
        if n is None:
            continue
        covered += 1
        if n < ctx.profile.min_text_size_norm:
            out.append(
                make_finding(
                    RULE_TINY_TEXT,
                    "typography",
                    Severity.MID,
                    f"文字过小（font_size_norm={n:.4f}）。",
                    0.6,
                    slide.index,
                    primary=e,
                    details={
                        "font_size_norm": round(n, 4),
                        "ratio_vs_min": round(n / ctx.profile.min_text_size_norm, 3),
                    },
                )
            )
    return RuleEvaluation(findings=out, covered_count=covered, eligible_count=len(eligible))


def flat_scale_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    fh = font_hierarchy(slide)
    title_el = slide.by_id(fh["title_id"]) if fh["title_id"] else None
    bodies = [e for e in slide.elements if e.role == "body" and e.has_text()]
    candidates = ([title_el] if title_el else []) + bodies
    eligible = len(candidates)
    covered = sum(1 for e in candidates if e.font_size_norm is not None)
    if fh["ratio"] is None or fh["ratio"] >= ctx.profile.flat_scale_ratio_min:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_FLAT_SCALE,
                "typography",
                Severity.LOW,
                f"标题/正文字号比 {fh['ratio']:.2f} 过平，缺乏层级。",
                0.5,
                slide.index,
                primary=title_el,
                details={"ratio": fh["ratio"]},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


RULES = [
    RuleSpec(rule_id=RULE_MANY_FAMILIES, dimension="typography", run=many_families_rule),
    RuleSpec(rule_id=RULE_TINY_TEXT, dimension="typography", run=tiny_text_rule),
    RuleSpec(rule_id=RULE_FLAT_SCALE, dimension="typography", run=flat_scale_rule),
]
