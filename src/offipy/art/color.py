"""颜色规则：对比度 / 强调色泛滥 / 无强调色。

rev2.2：前景色 foreground / run color；背景一律 effective_background（声明路径），
像素仅用于 declared_not_found 低置信提示，不参与对比度归因（冻结契约，#38）。
"""

from __future__ import annotations

from offipy.audit import Severity

from .features import effective_background, palette_features
from .models import ArtColor, ArtElement, ArtSlide
from .profiles import RULE_ACCENT_FLOOD, RULE_LOW_CONTRAST, RULE_NO_ACCENT
from .rules import RuleContext, RuleEvaluation, RuleSpec, make_finding


def _relative_luminance(c: ArtColor) -> float:
    def lin(v: float) -> float:
        v = v / 255.0
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = lin(c.r), lin(c.g), lin(c.b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(fg: ArtColor, bg: ArtColor) -> float:
    l1 = _relative_luminance(fg)
    l2 = _relative_luminance(bg)
    lo, hi = min(l1, l2), max(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def _effective_fg(el) -> list[ArtColor]:
    """文本前景：逐 run 取色，无 run 用元素 foreground。"""
    if el.runs:
        return [r.color for r in el.runs if r.color is not None]
    if el.foreground is not None:
        return [el.foreground]
    return []


def low_contrast_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    out = []
    eligible = [e for e in slide.elements if e.has_text()]
    covered = 0
    for e in eligible:
        fgs = _effective_fg(e)
        if not fgs:
            continue
        pe = e.pixel_evidence
        if (
            pe is not None
            and pe.method == "declared_not_found"
            and pe.foreground_match_ratio is not None
        ):
            # 声明前景在像素中未找到 → 低置信提示（<0.35 不驱动降级）
            covered += 1
            out.append(
                make_finding(
                    RULE_LOW_CONTRAST,
                    "color",
                    Severity.LOW,
                    f"声明前景色在页面像素中匹配率仅 "
                    f"{pe.foreground_match_ratio:.2f}，颜色可能异常。",
                    0.25,
                    slide.index,
                    primary=e,
                    details={"foreground_match_ratio": round(pe.foreground_match_ratio, 3)},
                    evidence_sources={"pixel"},
                    evidence_reliability=0.5,
                    evidence_method=pe.method,
                )
            )
            continue
        bg = effective_background(e, slide)
        if bg is None:
            continue  # 背景未知 → 无对比可判，covered 不加
        covered += 1
        for c in fgs:
            if c is None:
                continue
            if c.a < 1.0:
                c = c.with_alpha_over(bg)
            ratio = contrast_ratio(c, bg)
            if ratio < ctx.profile.min_contrast:
                sev = Severity.HIGH if ratio < 2.0 else Severity.MID
                out.append(
                    make_finding(
                        RULE_LOW_CONTRAST,
                        "color",
                        sev,
                        f"文本与背景对比度 {ratio:.2f} 低于 {ctx.profile.min_contrast}。",
                        0.95,
                        slide.index,
                        primary=e,  # rev2.1：实测对比度 → 高置信
                        details={"ratio": round(ratio, 3)},
                    )
                )
                break
    return RuleEvaluation(findings=out, covered_count=covered, eligible_count=len(eligible))


def _area_weighted_accent_ratio(slide: ArtSlide) -> tuple[float, object | None]:
    pal = palette_features(slide)
    ratio = pal["accent_ratio"]
    largest: ArtElement | None = None
    best = 0.0
    for e in slide.elements:
        c = e.foreground  # rev2.1：只认前景色
        if c is None or e.role in ("background", "container", "decoration"):
            continue
        sat = max(c.r, c.g, c.b) - min(c.r, c.g, c.b)
        if sat > 60 and e.area > best:
            best = e.area
            largest = e
    return ratio, largest


def _accent_rule_eval(slide: ArtSlide) -> tuple[int, int]:
    """评估范围=全部可见元素；covered=有前景色证据可判 accent 者。"""
    els = [e for e in slide.elements if e.role not in ("background", "container", "decoration")]
    eligible = len(els)
    covered = len([e for e in els if e.foreground is not None])
    return eligible, covered


def accent_flood_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    ratio, primary = _area_weighted_accent_ratio(slide)
    eligible, covered = _accent_rule_eval(slide)
    if covered == 0 or ratio <= ctx.profile.max_accent_ratio:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_ACCENT_FLOOD,
                "color",
                Severity.LOW,
                f"强调色占比 {ratio:.2f} 超过 {ctx.profile.max_accent_ratio}。",
                0.3,
                slide.index,
                primary=primary,
                details={"accent_ratio": ratio},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


def no_accent_rule(slide: ArtSlide, ctx: RuleContext) -> RuleEvaluation:
    ratio, _primary = _area_weighted_accent_ratio(slide)
    eligible, covered = _accent_rule_eval(slide)
    if covered == 0 or ratio != 0.0:
        return RuleEvaluation(covered_count=covered, eligible_count=eligible)
    return RuleEvaluation(
        findings=[
            make_finding(
                RULE_NO_ACCENT,
                "color",
                Severity.LOW,
                "页面完全没有强调色。",
                0.3,
                slide.index,
                details={"accent_ratio": 0.0},
            )
        ],
        covered_count=covered,
        eligible_count=eligible,
    )


RULES = [
    RuleSpec(rule_id=RULE_LOW_CONTRAST, dimension="color", run=low_contrast_rule),
    RuleSpec(
        rule_id=RULE_ACCENT_FLOOD, dimension="color", run=accent_flood_rule, experimental=True
    ),
    RuleSpec(rule_id=RULE_NO_ACCENT, dimension="color", run=no_accent_rule, experimental=True),
]
