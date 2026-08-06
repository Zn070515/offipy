"""Deck 级视觉一致性：标题位置/字号漂移、左边距漂移。

rev2.1：按 infer_slide_role 分组，每组 ≥3 页才判断；
过滤 background/container/decoration/page_number/footer。
"""

from __future__ import annotations

from offipy.audit import Severity

from .features import infer_slide_role
from .models import ArtFinding, ArtScene, ArtSlide
from .profiles import RULE_MARGIN_DRIFT, RULE_TITLE_DRIFT, ArtProfile
from .rules import make_finding

_FILTER_ROLES = {"background", "container", "decoration", "page_number", "footer"}


def _median(values: list[float]) -> float:
    ordered = sorted(values)
    n = len(ordered)
    if n == 0:
        return 0.0
    if n % 2 == 1:
        return ordered[n // 2]
    return (ordered[n // 2 - 1] + ordered[n // 2]) / 2.0


def _relative_drift(values: list[float]) -> float | None:
    if len(values) < 3:
        return None
    mean = sum(values) / len(values)
    if mean == 0:
        return None
    var = sum((v - mean) ** 2 for v in values) / len(values)
    return (var**0.5) / mean


def _title_of(slide: ArtSlide):
    return next((e for e in slide.elements if e.role == "title" and e.has_text()), None)


def _content_left_margin(slide: ArtSlide) -> float | None:
    vals = [e.x for e in slide.elements if e.area > 0 and e.role not in _FILTER_ROLES]
    return min(vals) if vals else None


def _title_drift_findings(
    group: list[ArtSlide], profile: ArtProfile, ref_slide: ArtSlide
) -> list[ArtFinding]:
    out: list[ArtFinding] = []
    title_ref = _title_of(ref_slide)
    if title_ref is None:
        return out
    xs = [_title_of(s).x for s in group if _title_of(s) is not None]
    x_drift = _relative_drift(xs) if len(xs) >= 3 else None
    sizes = [
        _title_of(s).font_size_norm
        for s in group
        if _title_of(s) is not None and _title_of(s).font_size_norm is not None
    ]
    size_drift = _relative_drift(sizes) if len(sizes) >= 3 else None
    drifted = False
    if x_drift is not None and x_drift > profile.title_drift_tol:
        drifted = True
    if size_drift is not None and size_drift > profile.title_drift_tol:
        drifted = True
    if drifted:
        out.append(
            make_finding(
                RULE_TITLE_DRIFT,
                "consistency",
                Severity.LOW,
                "标题位置或字号在页面间漂移，标题系统不一致。",
                0.5,
                ref_slide.index,
                primary=title_ref,
                details={
                    "x_drift": round(x_drift, 3) if x_drift else None,
                    "size_drift": round(size_drift, 3) if size_drift else None,
                    "slides": [s.index for s in group],
                },
            )
        )
    return out


def _margin_drift_findings(
    group: list[ArtSlide], profile: ArtProfile, ref_slide: ArtSlide
) -> list[ArtFinding]:
    margins: list[float] = []
    for s in group:
        m = _content_left_margin(s)
        if m is not None:
            margins.append(m)
    if len(margins) < 3:
        return []
    med = _median(margins)
    outliers = [m for m in margins if abs(m - med) > profile.margin_drift_tol]
    if not outliers:
        return []
    return [
        make_finding(
            RULE_MARGIN_DRIFT,
            "consistency",
            Severity.LOW,
            "页面左边距在页面间漂移。",
            0.45,
            ref_slide.index,
            details={"median": round(med, 3), "margins": [round(m, 3) for m in margins]},
        )
    ]


def _group_slides(scene: ArtScene) -> dict[str, list[ArtSlide]]:
    """按 infer_slide_role 分组；仅保留 ≥3 页的组。"""
    groups: dict[str, list[ArtSlide]] = {}
    for s in scene.slides:
        groups.setdefault(infer_slide_role(s), []).append(s)
    return {role: g for role, g in groups.items() if len(g) >= 3}


def assess_deck(scene: ArtScene, profile: ArtProfile) -> list[ArtFinding]:
    out: list[ArtFinding] = []
    for group in _group_slides(scene).values():
        ref_slide = group[0]
        out.extend(_title_drift_findings(group, profile, ref_slide))
        out.extend(_margin_drift_findings(group, profile, ref_slide))
    return out
