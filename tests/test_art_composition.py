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
    assert ev.findings[0].confidence <= 0.4  # experimental


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
    assert ev.findings[0].confidence <= 0.4


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
        "art.composition.background_like_area",
    }
    assert {rs.rule_id for rs in RULES if rs.experimental} == {
        "art.composition.off_balance",
        "art.composition.corner_cluster",
        "art.composition.background_like_area",
    }


def test_background_like_area_fires():
    from offipy.art.composition import background_like_area_rule
    from offipy.art.models import SlidePixelEvidence

    slide = make_slide(1, elements=[])
    slide.pixel_evidence = SlidePixelEvidence(
        background_like_ratio=0.9,
        background_confidence=0.9,
        background_uniformity=0.9,
    )
    ev = background_like_area_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.composition.background_like_area"
    assert ev.findings[0].confidence <= 0.4
    assert ev.findings[0].evidence_sources == frozenset({"pixel"})


def test_background_like_area_low_ratio_no_fire():
    from offipy.art.composition import background_like_area_rule
    from offipy.art.models import SlidePixelEvidence

    slide = make_slide(1, elements=[])
    slide.pixel_evidence = SlidePixelEvidence(
        background_like_ratio=0.3,
        background_confidence=0.9,
        background_uniformity=0.9,
    )
    assert background_like_area_rule(slide, _ctx(slide)).findings == []


def test_background_like_area_low_confidence_warns():
    from offipy.art.composition import background_like_area_rule
    from offipy.art.models import SlidePixelEvidence

    slide = make_slide(1, elements=[])
    slide.pixel_evidence = SlidePixelEvidence(
        background_like_ratio=0.9,
        background_confidence=0.4,  # 低于 0.7
        background_uniformity=0.9,
    )
    ev = background_like_area_rule(slide, _ctx(slide))
    assert ev.findings == []
    assert any(w.code == "art.pixel.background_low_confidence" for w in ev.warnings)


def test_background_like_area_high_occupancy_no_fire():
    from offipy.art.composition import background_like_area_rule
    from offipy.art.models import SlidePixelEvidence

    # 元素占满页面 → union_area_ratio > 0.5 → 不提示、不警告
    slide = make_slide(
        1,
        elements=[
            ArtElement(
                element_id="full",
                kind="shape",
                role="body",
                x=0.0,
                y=0.0,
                width=1.0,
                height=0.9,
                slide_index=1,
            )
        ],
    )
    slide.pixel_evidence = SlidePixelEvidence(
        background_like_ratio=0.9,
        background_confidence=0.9,
        background_uniformity=0.9,
    )
    ev = background_like_area_rule(slide, _ctx(slide))
    assert ev.findings == []
    assert ev.warnings == []


def test_background_like_area_full_bleed_image_no_fire():
    from offipy.art.composition import background_like_area_rule
    from offipy.art.models import SlidePixelEvidence

    slide = make_slide(
        1,
        elements=[
            ArtElement(
                element_id="img",
                kind="image",
                role="image",
                x=0.0,
                y=0.0,
                width=1.0,
                height=1.0,
                slide_index=1,
            )
        ],
    )
    slide.pixel_evidence = SlidePixelEvidence(
        background_like_ratio=0.9,
        background_confidence=0.9,
        background_uniformity=0.9,
    )
    assert background_like_area_rule(slide, _ctx(slide)).findings == []


def test_background_like_area_full_bleed_background_image_no_fire():
    """is_background 全幅图片被 density 排除 → occupancy 低 → _full_bleed_image 分支真正触发。"""
    from offipy.art.composition import background_like_area_rule
    from offipy.art.models import SlidePixelEvidence

    slide = make_slide(
        1,
        elements=[
            ArtElement(
                element_id="bg_img",
                kind="image",
                role="image",
                x=0.0,
                y=0.0,
                width=1.0,
                height=1.0,
                slide_index=1,
                is_background=True,
            )
        ],
    )
    slide.pixel_evidence = SlidePixelEvidence(
        background_like_ratio=0.9,
        background_confidence=0.9,
        background_uniformity=0.9,
    )
    assert background_like_area_rule(slide, _ctx(slide)).findings == []
