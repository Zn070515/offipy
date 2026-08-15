from art_helpers import make_image_element, make_slide
from offipy.art.features import compute_features
from offipy.art.media import (
    RULES,
    distorted_image_rule,
    mixed_image_sizes_rule,
    tiny_image_rule,
)
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


def test_distorted_image_fires():
    # 物理比 4:1（960×240），解码比 4:3 → 漂移大
    slide = make_slide(
        1,
        elements=[
            make_image_element("i", w=0.5, h=0.125, decoded_width=800.0, decoded_height=600.0),
        ],
    )
    ev = distorted_image_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.media.distorted_image"
    assert ev.findings[0].severity == Severity.MID


def test_distorted_image_ok():
    # 物理 (0.45×1920)/(0.6×1080) = 864/648 = 1.333 ≈ 解码 800/600=1.333 → 不触发
    slide = make_slide(
        1,
        elements=[
            make_image_element("i", w=0.45, h=0.6, decoded_width=800.0, decoded_height=600.0),
        ],
    )
    assert distorted_image_rule(slide, _ctx(slide)).findings == []


def test_tiny_image_fires():
    slide = make_slide(
        1,
        elements=[
            make_image_element("i", w=0.01, h=0.01),
        ],
    )
    ev = tiny_image_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.media.tiny_image"


def test_tiny_image_details_area_ratio():
    # 面积 0.04×0.04=0.0016 < balanced min_image_area 0.0025 → 触发；
    # details["area_ratio"] 必须是面积占比按 3 位小数取整（与 distorted_image 的 details 模式对齐）
    el = make_image_element("i", w=0.04, h=0.04)
    slide = make_slide(1, elements=[el])
    ev = tiny_image_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].details["area_ratio"] == 0.002  # round(0.0016, 3)：钉死 3 位小数约定


def test_mixed_image_sizes_fires():
    slide = make_slide(
        1,
        elements=[
            make_image_element("a", w=0.4, h=0.3),
            make_image_element("b", w=0.5, h=0.375),
            make_image_element("c", w=0.02, h=0.015),  # 面积差一个量级
        ],
    )
    ev = mixed_image_sizes_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.media.mixed_image_sizes"
    assert ev.findings[0].primary is not None  # 最大图


def test_mixed_image_sizes_needs_three():
    slide = make_slide(
        1,
        elements=[
            make_image_element("a", w=0.4, h=0.3),
            make_image_element("b", w=0.5, h=0.375),
        ],
    )
    assert mixed_image_sizes_rule(slide, _ctx(slide)).findings == []


def test_media_rules_are_rule_specs():
    assert all(isinstance(rs, RuleSpec) for rs in RULES)
    assert {rs.rule_id for rs in RULES} == {
        "art.media.distorted_image",
        "art.media.tiny_image",
        "art.media.mixed_image_sizes",
    }
