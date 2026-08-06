from art_helpers import make_slide
from offipy.art.composition import (
    RULES,
    corner_cluster_rule,
    off_balance_rule,
    spacing_drift_rule,
)
from offipy.art.features import compute_features, spacing_features
from offipy.art.models import ArtElement
from offipy.art.profiles import get_profile
from offipy.art.rules import RuleContext, RuleSpec


def _ctx(slide, profile="balanced"):
    return RuleContext(
        profile=get_profile(profile),
        slide=slide,
        slide_index=slide.index,
        features=compute_features(slide),
        deck=__import__("offipy.art.models", fromlist=["ArtScene"]).ArtScene(slides=[slide]),
    )


def _el(element_id, x, y, w, h, role="body", kind="shape"):
    return ArtElement(
        element_id=element_id, kind=kind, role=role, x=x, y=y, width=w, height=h, slide_index=1
    )


def test_off_balance_fires():
    # 三个元素全挤在左侧 → 质量失衡
    slide = make_slide(
        1,
        elements=[
            _el("a", 0.0, 0.2, 0.2, 0.4, kind="text", role="body"),
            _el("b", 0.05, 0.3, 0.2, 0.4, kind="text", role="body"),
            _el("c", 0.1, 0.4, 0.2, 0.4, kind="text", role="body"),
        ],
    )
    ev = off_balance_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.composition.off_balance"
    assert ev.findings[0].confidence <= 0.3  # experimental


def test_off_balance_balanced_no_fire():
    slide = make_slide(
        1,
        elements=[
            _el("a", 0.1, 0.3, 0.2, 0.2, kind="text", role="body"),
            _el("b", 0.5, 0.3, 0.2, 0.2, kind="text", role="body"),
            _el("c", 0.3, 0.1, 0.2, 0.2, kind="text", role="body"),
        ],
    )
    assert off_balance_rule(slide, _ctx(slide)).findings == []


def test_corner_cluster_fires():
    slide = make_slide(
        1,
        elements=[
            _el("a", 0.0, 0.0, 0.3, 0.3),
            _el("b", 0.05, 0.05, 0.2, 0.2),
            _el("c", 0.02, 0.02, 0.15, 0.15),
            _el("d", 0.6, 0.6, 0.1, 0.1),
        ],
    )
    ev = corner_cluster_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.composition.corner_cluster"
    assert ev.findings[0].confidence <= 0.3


def test_spacing_drift_fires():
    slide = make_slide(
        1,
        elements=[
            _el("a", 0.0, 0.0, 0.2, 0.1),
            _el("b", 0.4, 0.0, 0.2, 0.1),
            _el("c", 0.8, 0.0, 0.2, 0.1),
            _el("d", 1.5, 0.0, 0.2, 0.1),  # 与 c 间距 0.5，偏离中位 0.2
        ],
    )
    sp = spacing_features(slide.elements)
    assert sp["horizontal"]["drift_count"] >= 1
    ev = spacing_drift_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1


def test_spacing_drift_ok_regular():
    slide = make_slide(
        1,
        elements=[
            _el("a", 0.0, 0.0, 0.2, 0.1),
            _el("b", 0.4, 0.0, 0.2, 0.1),
            _el("c", 0.8, 0.0, 0.2, 0.1),
        ],
    )
    assert spacing_drift_rule(slide, _ctx(slide)).findings == []


def test_composition_rules_are_rule_specs():
    assert all(isinstance(rs, RuleSpec) for rs in RULES)
    assert {rs.rule_id for rs in RULES} == {
        "art.composition.off_balance",
        "art.composition.corner_cluster",
        "art.composition.spacing_drift",
    }
