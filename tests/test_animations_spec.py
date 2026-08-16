"""animations/spec.py：声明归一化（显式优先 + 约定回退 + 校验 + 告警）。"""

import pytest

from offipy.animations.spec import (
    EFFECT_CATALOG,
    AnimationSpec,
    ParsedAnimation,
    TransitionSpec,
    parse_declaration,
)
from offipy.exceptions import InvalidArgumentError


def test_parse_explicit_fly_in_left():
    spec = parse_declaration(
        {"anim": "fly_in", "dir": "left", "trigger": "click", "dur": "0.8"},
        {},
    )
    assert spec == ParsedAnimation(
        effect="fly_in", direction="left", trigger="click", duration=0.8, delay=0.0
    )


def test_parse_explicit_defaults():
    spec = parse_declaration({"anim": "fade"}, {})
    assert spec.effect == "fade"
    assert spec.direction == "bottom"  # direction 对 fade 忽略但仍给默认
    assert spec.trigger == "click"
    assert spec.duration == pytest.approx(0.5)
    assert spec.delay == 0.0


def test_parse_after_trigger_keeps_delay():
    spec = parse_declaration({"anim": "fade", "trigger": "after", "delay": "0.2"}, {})
    assert spec.trigger == "after"
    assert spec.delay == pytest.approx(0.2)


def test_parse_explicit_unknown_effect_returns_none():
    assert parse_declaration({"anim": "wobble"}, {}) is None


def test_parse_fallback_fade_via_data_anim():
    spec = parse_declaration({"dataAnim": "fade"}, {})
    assert spec is not None and spec.effect == "fade"


def test_parse_fallback_fade_via_data_aos():
    spec = parse_declaration({"dataAos": "fade"}, {})
    assert spec is not None and spec.effect == "fade"


def test_parse_fallback_fade_via_class():
    spec = parse_declaration({"fadeIn": True}, {})
    assert spec is not None and spec.effect == "fade"


def test_parse_fallback_fade_up_float():
    spec = parse_declaration({"dataAnim": "fade-up"}, {})
    assert spec is not None and spec.effect == "float_up"


def test_parse_fallback_direction_mappings():
    assert parse_declaration({"dataAnim": "fade-left"}, {}).direction == "left"
    assert parse_declaration({"dataAnim": "fade-right"}, {}).direction == "right"
    assert parse_declaration({"dataAnim": "fade-down"}, {}).direction == "top"
    assert parse_declaration({"dataAos": "zoom-in"}, {}).effect == "zoom_in"


def test_parse_fallback_unmapped_skipped():
    # flip / slide-up / fade-up-right 等未列入 → None，不产动画
    assert parse_declaration({"dataAnim": "flip"}, {}) is None
    assert parse_declaration({"dataAnim": "slide-up"}, {}) is None
    assert parse_declaration({"dataAnim": "fade-up-right"}, {}) is None


def test_parse_bbox_replacement_record_skipped():
    # kind='asset'（data-asset/data-icon/data-primitive）与 mermaid/drawio className
    assert parse_declaration({"anim": "fade"}, {"kind": "asset"}) is None
    assert parse_declaration({"anim": "fade"}, {"className": "mermaid"}) is None
    assert parse_declaration({"anim": "fade"}, {"className": "drawio"}) is None


def test_parse_explicit_invalid_direction_skipped():
    assert parse_declaration({"anim": "fly_in", "dir": "northwest"}, {}) is None


def test_parse_explicit_invalid_trigger_skipped():
    assert parse_declaration({"anim": "fade", "trigger": "hover"}, {}) is None


def test_parse_explicit_non_numeric_duration_skipped():
    assert parse_declaration({"anim": "fade", "dur": "abc"}, {}) is None


def test_effect_catalog_complete():
    assert set(EFFECT_CATALOG) == {"fade", "float_up", "fly_in", "wipe", "zoom_in", "grow"}


def test_transition_spec_validation():
    t = TransitionSpec(slide=2, kind="push", speed="medium")
    assert t.kind == "push" and t.speed == "medium"


def test_animation_spec_invalid_slide_raises():
    # trigger/duration/delay 有默认值，slide/target/effect 才是真正的必填；
    # slide<1 触发 __post_init__ 校验
    with pytest.raises(InvalidArgumentError):
        AnimationSpec(slide=0, target="t", effect="fade")
