import pytest

from art_helpers import make_image_element, make_slide, make_text_element
from offipy.art import encode_features
from offipy.art.features import compute_features
from offipy.art.hierarchy import RULES, no_focus_rule, title_too_small_rule
from offipy.art.profiles import get_profile
from offipy.art.rules import RuleContext, RuleSpec
from offipy.audit import Severity


def _ctx(slide, profile="balanced"):
    return RuleContext(
        profile=get_profile(profile),
        slide=slide,
        slide_index=slide.index,
        features=compute_features(slide),
        deck=__import__("offipy.art.models", fromlist=["ArtScene"]).ArtScene(slides=[slide]),
    )


def test_title_too_small_fires():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=20.0, role="title"),  # 20/1080<0.03
            make_text_element("b", "Body", y=0.2, font_size=24.0, role="body"),
        ],
    )
    ev = title_too_small_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.hierarchy.title_too_small"
    assert ev.findings[0].severity == Severity.MID
    assert ev.findings[0].confidence == 0.55


def test_title_too_small_ok():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=54.0, role="title"),
            make_text_element("b", "Body", y=0.2, font_size=24.0, role="body"),
        ],
    )
    assert title_too_small_rule(slide, _ctx(slide)).findings == []


def test_title_too_small_details():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=20.0, role="title"),  # 20/1080 < 0.03
        ],
    )
    ev = title_too_small_rule(slide, _ctx(slide))
    f = ev.findings[0]
    n = 20.0 / 1080.0
    assert "font_size_norm" in f.details
    assert "ratio_vs_min" in f.details
    assert f.details["ratio_vs_min"] == pytest.approx(n / 0.03, rel=1e-3)


def test_no_focus_fires_when_flat():
    slide = make_slide(
        1,
        elements=[
            make_text_element("a", "A", font_size=24.0),
            make_text_element("b", "B", y=0.2, font_size=24.0),
            make_text_element("c", "C", y=0.3, font_size=24.0),
        ],
    )
    ev = no_focus_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.hierarchy.no_focus"
    assert ev.findings[0].confidence <= 0.4  # experimental


def test_no_focus_details_focus_ratio():
    slide = make_slide(
        1,
        elements=[
            make_text_element("a", "A", font_size=24.0),
            make_text_element("b", "B", y=0.2, font_size=24.0),
            make_text_element("c", "C", y=0.3, font_size=24.0),
        ],
    )
    ev = no_focus_rule(slide, _ctx(slide))
    f = ev.findings[0]
    assert "focus_ratio" in f.details


def test_no_focus_details_focus_ratio_none_when_no_sizes():
    # 3 个无字号的 image：focus 特征拿不到 ratio → details 显式 None，
    # encode_features 对 measure 特征回退 missing_default 0.0
    slide = make_slide(
        1,
        elements=[
            make_image_element("a", x=0.05, y=0.05),
            make_image_element("b", x=0.55, y=0.05),
            make_image_element("c", x=0.05, y=0.55),
        ],
    )
    ev = no_focus_rule(slide, _ctx(slide))
    f = ev.findings[0]
    assert f.details["focus_ratio"] is None
    enc = encode_features(f, slide, deck=None, profile="balanced")
    assert enc["measure.art.hierarchy.no_focus.focus_ratio"] == 0.0


def test_no_focus_ok_when_dominant():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=72.0, role="title"),
            make_text_element("b", "Body", y=0.2, font_size=20.0, role="body"),
            make_text_element("c", "Cap", y=0.3, font_size=20.0, role="caption"),
        ],
    )
    assert no_focus_rule(slide, _ctx(slide)).findings == []


def test_no_focus_needs_three_elements():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=24.0),
            make_text_element("b", "Body", y=0.2, font_size=24.0),
        ],
    )
    ev = no_focus_rule(slide, _ctx(slide))
    assert ev.findings == []
    assert ev.eligible_count == 0  # 元素不足 → 不适用，不进 coverage
    assert ev.covered_count == 0


def test_hierarchy_rules_are_rule_specs():
    assert all(isinstance(rs, RuleSpec) for rs in RULES)
    assert {rs.rule_id for rs in RULES} == {
        "art.hierarchy.no_focus",
        "art.hierarchy.title_too_small",
    }
