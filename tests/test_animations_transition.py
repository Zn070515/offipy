"""animations/transition.py：<p:transition> OOXML 构建（4 种 classic 过渡）。"""

import pytest

from offipy.animations.transition import build_transition
from offipy.exceptions import InvalidArgumentError


def _qn(tag):
    return tag.split("}")[-1]


def test_transition_fade():
    el = build_transition("fade", "medium")
    assert _qn(el.tag) == "transition"
    assert el.get("spd") == "med"
    child = next(iter(el))
    assert _qn(child.tag) == "fade"


def test_transition_wipe_push_cover():
    for kind, has_dir in (("wipe", True), ("push", True), ("cover", True)):
        el = build_transition(kind, "slow")
        assert el.get("spd") == "slow"
        child = next(iter(el))
        assert _qn(child.tag) == kind
        if has_dir:
            assert child.get("dir") == "l"


def test_transition_speed_values():
    assert build_transition("fade", "slow").get("spd") == "slow"
    assert build_transition("fade", "medium").get("spd") == "med"
    assert build_transition("fade", "fast").get("spd") == "fast"


def test_transition_unknown_kind_raises():
    with pytest.raises(InvalidArgumentError):
        build_transition("flip")


def test_transition_unknown_speed_raises():
    with pytest.raises(InvalidArgumentError):
        build_transition("fade", speed="turbo")
