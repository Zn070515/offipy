"""animations/apply.py：独立 API——按形状名定位 spid → 注入 timing/transition → 报告。"""

import pytest

from offipy.animations.apply import apply_animations, apply_transitions
from offipy.animations.spec import AnimationSpec, TransitionSpec
from offipy.exceptions import InvalidArgumentError


def _build_pptx(tmp_path, shapes, *, existing_timing=False):
    """造一个 1 页 .pptx，shapes=[(name,)]。返回路径。"""
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for i, (name,) in enumerate(shapes):
        sp = slide.shapes.add_shape(1, Inches(1), Inches(1 + i), Inches(2), Inches(0.6))
        sp.name = name
    if existing_timing:
        # 给 slide 塞一个假 <p:timing>（构造字符串插入）
        t = slide._element.makeelement(qn("p:timing"), {})
        slide._element.append(t)
    p = tmp_path / "a.pptx"
    prs.save(str(p))
    return str(p)


def _qn(tag):
    return tag.split("}")[-1]


def test_apply_animations_injects_timing(tmp_path):
    p = _build_pptx(tmp_path, [("title",)])
    report = apply_animations(
        p,
        animations=[AnimationSpec(slide=1, target="title", effect="fade")],
    )
    assert report["animations_applied"] == 1
    assert report["transitions_applied"] == 0
    assert report["unmatched"] == []
    from pptx import Presentation

    sld = Presentation(p).slides[0]
    timing = sld._element.findall(".//{*}timing")
    assert len(timing) == 1


def test_apply_animations_target_not_found_reported_not_fatal(tmp_path):
    p = _build_pptx(tmp_path, [("title",)])
    report = apply_animations(
        p,
        animations=[AnimationSpec(slide=1, target="missing", effect="fade")],
        raise_on_all_unmatched=False,
    )
    assert report["unmatched"] == [{"slide": 1, "target": "missing"}]
    assert report["animations_applied"] == 0


def test_apply_animations_all_unmatched_raises(tmp_path):
    p = _build_pptx(tmp_path, [("title",)])
    with pytest.raises(InvalidArgumentError):
        apply_animations(
            p,
            animations=[
                AnimationSpec(slide=1, target="a", effect="fade"),
                AnimationSpec(slide=1, target="b", effect="fade"),
            ],
        )


def test_apply_animations_transition_only(tmp_path):
    p = _build_pptx(tmp_path, [("title",)])
    report = apply_animations(
        p,
        transitions=[TransitionSpec(slide=1, kind="push", speed="medium")],
    )
    assert report["transitions_applied"] == 1
    assert report["animations_applied"] == 0
    from pptx import Presentation

    sld = Presentation(p).slides[0]
    transitions = sld._element.findall(".//{*}transition")
    assert len(transitions) == 1


def test_apply_animations_order_transition_before_timing(tmp_path):
    p = _build_pptx(tmp_path, [("title",)])
    apply_animations(
        p,
        animations=[AnimationSpec(slide=1, target="title", effect="fade")],
        transitions=[TransitionSpec(slide=1, kind="fade", speed="fast")],
    )
    from pptx import Presentation

    sld = Presentation(p).slides[0]
    names = [_qn(c.tag) for c in sld._element]
    assert names.index("transition") < names.index("timing")
    # 且都在 clrMapOvr 之后
    assert names.index("cSld") < names.index("transition")


def test_apply_animations_idempotency_raises(tmp_path):
    p = _build_pptx(tmp_path, [("title",)], existing_timing=True)
    with pytest.raises(InvalidArgumentError, match="已有动画"):
        apply_animations(
            p,
            animations=[AnimationSpec(slide=1, target="title", effect="fade")],
        )


def test_apply_transitions_wrapper(tmp_path):
    p = _build_pptx(tmp_path, [("title",)])
    report = apply_transitions(p, [TransitionSpec(slide=1, kind="cover", speed="slow")])
    assert report["transitions_applied"] == 1
