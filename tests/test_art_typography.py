from art_helpers import make_slide, make_text_element
from offipy.art.features import compute_features
from offipy.art.models import ArtElement, ArtTextRun
from offipy.art.profiles import get_profile
from offipy.art.rules import RuleContext
from offipy.art.typography import (
    RULES,
    flat_scale_rule,
    many_families_rule,
    tiny_text_rule,
)
from offipy.audit import Severity


def _ctx(slide, profile="balanced"):
    return RuleContext(
        profile=get_profile(profile),
        slide=slide,
        slide_index=slide.index,
        features=compute_features(slide),
        deck=__import__("offipy.art.models", fromlist=["ArtScene"]).ArtScene(slides=[slide]),
    )


def _families(*families):
    return ArtElement(
        element_id="t",
        kind="text",
        role="body",
        x=0.1,
        y=0.1,
        width=0.5,
        height=0.08,
        slide_index=1,
        text="x",
        runs=[ArtTextRun(text="x", font_family=f) for f in families],
        font_size=24.0,
        font_size_unit="px",
    )


def test_many_families_fires_over_three():
    slide = make_slide(
        1,
        elements=[
            _families("A", "B", "C", "D"),
            _families("A", "B", "C"),
            make_text_element("t2", "Body", font_size=24.0),
        ],
    )
    ev = many_families_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.typography.many_families"
    assert ev.findings[0].severity == Severity.MID
    assert ev.findings[0].primary is not None


def test_many_families_ok_at_three():
    slide = make_slide(
        1,
        elements=[
            _families("A", "B", "C"),
            _families("A", "B"),
            make_text_element("t2", "Body", font_size=24.0),
        ],
    )
    assert many_families_rule(slide, _ctx(slide)).findings == []


def test_tiny_text_fires_below_norm():
    slide = make_slide(
        1,
        elements=[
            make_text_element("small", "Small", font_size=10.0),  # 10/1080 < 0.015
            make_text_element("big", "Big", font_size=30.0),
        ],
    )
    ev = tiny_text_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.typography.tiny_text"


def test_flat_scale_fires_low_ratio():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=30.0, role="title"),
            make_text_element("b", "Body", y=0.2, font_size=27.0, role="body"),
        ],
    )
    ev = flat_scale_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.typography.flat_scale"
    assert ev.findings[0].severity == Severity.LOW


def test_flat_scale_ok_with_clear_ratio():
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Title", font_size=54.0, role="title"),
            make_text_element("b", "Body", y=0.2, font_size=24.0, role="body"),
        ],
    )
    assert flat_scale_rule(slide, _ctx(slide)).findings == []


def test_typography_coverage_drops_without_font_evidence():
    # 无 font_size_norm（如 pptx 源）→ covered 为 0，覆盖降级
    slide = make_slide(
        1,
        elements=[
            make_text_element("a", "A", font_size=None),
            make_text_element("b", "B", font_size=None),
        ],
    )
    ev = tiny_text_rule(slide, _ctx(slide))
    assert ev.eligible_count == 2
    assert ev.covered_count == 0


def test_rules_are_rule_specs():
    from offipy.art.rules import RuleSpec

    assert all(isinstance(rs, RuleSpec) for rs in RULES)
    assert {rs.rule_id for rs in RULES} == {
        "art.typography.many_families",
        "art.typography.tiny_text",
        "art.typography.flat_scale",
    }
