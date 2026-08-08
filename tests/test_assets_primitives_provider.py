"""A5 Task 1 — primitives provider contracts.

Strict per-primitive schemas, exact min/max lengths, enum case policy,
integer syntax boundary, process-arrow list rules, unknown/forbidden params,
and deterministic search order.
"""

from __future__ import annotations

import pytest

from offipy.assets.model import AssetRef, AssetRequest, NativeShapePayload
from offipy.assets.providers.primitives import (
    _PRIMITIVE_ORDER,
    PrimitivesProvider,
)
from offipy.assets.registry import get_default_registry
from offipy.exceptions import InvalidArgumentError


def _req(name: str, params: tuple[tuple[str, str], ...] = ()) -> AssetRequest:
    return AssetRequest(ref=AssetRef("primitives", "primitive", name), params=params)


def _resolve(name: str, params: tuple[tuple[str, str], ...] = ()) -> NativeShapePayload:
    resolved = PrimitivesProvider().resolve(_req(name, params))
    assert isinstance(resolved.payload, NativeShapePayload)
    return resolved.payload


# -- common params ---------------------------------------------------------

# required params for primitives that have them (device-frame), so common-param
# tests can exercise every primitive without tripping required-param validation.
_REQUIRED_BASE: dict[str, tuple[tuple[str, str], ...]] = {
    "quote-mark": (("text", "hi"),),
    "section-number": (("number", "1"),),
    "label-pill": (("text", "hi"),),
    "metric-badge": (("value", "1"),),
    "timeline-node": (),
    "process-arrow": (("steps", "a,b"),),
    "device-frame": (("device", "phone"),),
    "browser-mockup": (),
}


@pytest.mark.parametrize("name", _PRIMITIVE_ORDER)
def test_every_primitive_accepts_accent_and_fill(name: str) -> None:
    payload = _resolve(name, _REQUIRED_BASE[name] + (("accent", "#123456"), ("fill", "surface")))
    params = dict(payload.params)
    assert params["accent"] == "#123456"
    assert params["fill"] == "surface"


@pytest.mark.parametrize("name", _PRIMITIVE_ORDER)
def test_accent_and_fill_defaults(name: str) -> None:
    payload = _resolve(name, _REQUIRED_BASE[name])
    params = dict(payload.params)
    assert params["accent"] == "accent"
    assert params["fill"] in ("transparent", "accent", "surface")


# -- quote-mark ------------------------------------------------------------


def test_quote_mark_requires_text() -> None:
    with pytest.raises(InvalidArgumentError, match="text"):
        _resolve("quote-mark")


def test_quote_mark_text_trims() -> None:
    payload = _resolve("quote-mark", (("text", "  Hello world  "),))
    assert dict(payload.params)["text"] == "Hello world"


def test_quote_mark_text_max_length() -> None:
    ok = _resolve("quote-mark", (("text", "x" * 240),))
    assert dict(ok.params)["text"] == "x" * 240
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("quote-mark", (("text", "x" * 241),))


def test_quote_mark_text_whitespace_only_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="non-empty"):
        _resolve("quote-mark", (("text", "   "),))


# -- section-number --------------------------------------------------------


@pytest.mark.parametrize("value", ["0", "1", "9999"])
def test_section_number_boundary_accepted(value: str) -> None:
    payload = _resolve("section-number", (("number", value),))
    assert dict(payload.params)["number"] == value


@pytest.mark.parametrize("value", ["-1", "10000", "1.5", "abc", "007a"])
def test_section_number_invalid_rejected(value: str) -> None:
    with pytest.raises(InvalidArgumentError):
        _resolve("section-number", (("number", value),))


def test_section_number_label_optional() -> None:
    assert "label" not in dict(_resolve("section-number", (("number", "3"),)).params)
    payload = _resolve("section-number", (("number", "3"), ("label", "Intro")))
    assert dict(payload.params)["label"] == "Intro"


def test_section_number_label_max_length() -> None:
    _resolve("section-number", (("number", "3"), ("label", "x" * 120)))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("section-number", (("number", "3"), ("label", "x" * 121)))


# -- label-pill ------------------------------------------------------------


def test_label_pill_requires_text() -> None:
    with pytest.raises(InvalidArgumentError, match="text"):
        _resolve("label-pill")


def test_label_pill_text_max_length() -> None:
    _resolve("label-pill", (("text", "x" * 120),))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("label-pill", (("text", "x" * 121),))


# -- metric-badge ----------------------------------------------------------


def test_metric_badge_requires_value() -> None:
    with pytest.raises(InvalidArgumentError, match="value"):
        _resolve("metric-badge")


def test_metric_badge_full_params() -> None:
    payload = _resolve(
        "metric-badge",
        (("value", "24%"), ("label", "YoY"), ("delta", "+3.2%")),
    )
    params = dict(payload.params)
    assert params["value"] == "24%"
    assert params["label"] == "YoY"
    assert params["delta"] == "+3.2%"


def test_metric_badge_lengths() -> None:
    _resolve("metric-badge", (("value", "x" * 80), ("label", "y" * 120), ("delta", "z" * 40)))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("metric-badge", (("value", "x" * 81),))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("metric-badge", (("value", "v"), ("label", "y" * 121)))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("metric-badge", (("value", "v"), ("delta", "z" * 41)))


# -- timeline-node ---------------------------------------------------------


def test_timeline_node_phase_default() -> None:
    payload = _resolve("timeline-node")
    assert dict(payload.params)["phase"] == "current"


@pytest.mark.parametrize("phase", ["past", "current", "future"])
def test_timeline_node_phase_accepted(phase: str) -> None:
    payload = _resolve("timeline-node", (("phase", phase),))
    assert dict(payload.params)["phase"] == phase


def test_timeline_node_phase_unknown_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="phase"):
        _resolve("timeline-node", (("phase", "now"),))


def test_timeline_node_label_optional() -> None:
    assert "label" not in dict(_resolve("timeline-node").params)
    payload = _resolve("timeline-node", (("label", "Launch"),))
    assert dict(payload.params)["label"] == "Launch"


# -- process-arrow ---------------------------------------------------------


def test_process_arrow_requires_steps() -> None:
    with pytest.raises(InvalidArgumentError, match="steps"):
        _resolve("process-arrow")


def test_process_arrow_steps_min_max() -> None:
    payload = _resolve("process-arrow", (("steps", "a,b"),))
    assert dict(payload.params)["steps"] == "a,b"
    payload = _resolve("process-arrow", (("steps", "a,b,c,d,e,f,g,h"),))
    assert len(dict(payload.params)["steps"].split(",")) == 8
    with pytest.raises(InvalidArgumentError, match="2..8"):
        _resolve("process-arrow", (("steps", "a"),))
    with pytest.raises(InvalidArgumentError, match="2..8"):
        _resolve("process-arrow", (("steps", "a,b,c,d,e,f,g,h,i"),))


def test_process_arrow_empty_item_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="empty"):
        _resolve("process-arrow", (("steps", "a,,b"),))


def test_process_arrow_steps_trim_and_max() -> None:
    payload = _resolve("process-arrow", (("steps", " a , b "),))
    assert dict(payload.params)["steps"] == "a,b"
    _resolve("process-arrow", (("steps", "x" * 80 + "," + "y" * 80),))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("process-arrow", (("steps", "x" * 81 + ",y"),))


def test_process_arrow_direction_default_and_enum() -> None:
    payload = _resolve("process-arrow", (("steps", "a,b"),))
    assert dict(payload.params)["direction"] == "horizontal"
    payload = _resolve("process-arrow", (("steps", "a,b"), ("direction", "vertical")))
    assert dict(payload.params)["direction"] == "vertical"
    with pytest.raises(InvalidArgumentError, match="direction"):
        _resolve("process-arrow", (("steps", "a,b"), ("direction", "diagonal")))


# -- device-frame ----------------------------------------------------------


def test_device_frame_requires_device() -> None:
    with pytest.raises(InvalidArgumentError, match="device"):
        _resolve("device-frame")


@pytest.mark.parametrize("device", ["phone", "tablet", "desktop"])
def test_device_frame_devices(device: str) -> None:
    payload = _resolve("device-frame", (("device", device),))
    assert dict(payload.params)["device"] == device


def test_device_frame_unknown_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="device"):
        _resolve("device-frame", (("device", "watch"),))


# -- browser-mockup --------------------------------------------------------


def test_browser_mockup_all_optional() -> None:
    params = dict(_resolve("browser-mockup").params)
    assert "title" not in params
    assert "url" not in params


def test_browser_mockup_title_url() -> None:
    payload = _resolve("browser-mockup", (("title", "Dashboard"), ("url", "https://x")))
    params = dict(payload.params)
    assert params["title"] == "Dashboard"
    assert params["url"] == "https://x"


def test_browser_mockup_lengths() -> None:
    _resolve("browser-mockup", (("title", "x" * 120), ("url", "y" * 240)))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("browser-mockup", (("title", "x" * 121),))
    with pytest.raises(InvalidArgumentError, match="max length"):
        _resolve("browser-mockup", (("url", "y" * 241),))


# -- unknown / forbidden params --------------------------------------------


@pytest.mark.parametrize("name", _PRIMITIVE_ORDER)
def test_unknown_param_rejected(name: str) -> None:
    with pytest.raises(InvalidArgumentError, match="unknown param"):
        _resolve(name, (("bogus", "x"),))


@pytest.mark.parametrize("forbidden", ["screenshot", "src", "image"])
def test_forbidden_params_explicit_v014(forbidden: str) -> None:
    with pytest.raises(InvalidArgumentError, match="not supported in v0.14"):
        _resolve("browser-mockup", ((forbidden, "x"),))


def test_bad_color_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="color"):
        _resolve("label-pill", (("text", "hi"), ("accent", "zzz")))


# -- control characters ----------------------------------------------------


@pytest.mark.parametrize("bad", ["\x00", "\x07", "\x1b", "\x7f", "\x85"])
def test_text_param_rejects_control_chars(bad: str) -> None:
    # 控制字符会污染 OOXML run 文本 / 可走私转义字节进 deck，一律拒绝
    with pytest.raises(InvalidArgumentError, match="control character"):
        _resolve("label-pill", (("text", f"hi{bad}there"),))


def test_list_item_rejects_control_chars() -> None:
    with pytest.raises(InvalidArgumentError, match="control character"):
        _resolve("process-arrow", (("steps", "a,hi\x1bthere"),))


def test_tab_and_newline_allowed_in_text() -> None:
    # \t \n \r 是合法排版空白，保留（仅拒绝其它控制字符）
    payload = _resolve("quote-mark", (("text", "line1\nline2\ttab"),))
    assert dict(payload.params)["text"] == "line1\nline2\ttab"


# -- provider metadata / search --------------------------------------------


def test_provider_metadata() -> None:
    provider = PrimitivesProvider()
    assert provider.provider_id == "primitives"
    assert provider.kinds == frozenset({"primitive"})
    assert provider.provider_meta.license == "MIT"
    assert provider.provider_meta.first_party is True
    assert provider.provider_meta.redistributable is True


def test_search_order_stable() -> None:
    provider = PrimitivesProvider()
    names = [m.ref.name for m in provider.search("", limit=100)]
    assert names == list(_PRIMITIVE_ORDER)


def test_search_kind_filter() -> None:
    provider = PrimitivesProvider()
    assert provider.search("quote", kind="pattern") == []
    names = [m.ref.name for m in provider.search("quote")]
    assert names == ["quote-mark"]


def test_search_limit() -> None:
    provider = PrimitivesProvider()
    assert len(provider.search("", limit=3)) == 3


def test_registry_includes_primitives() -> None:
    registry = get_default_registry()
    resolved = registry.resolve(
        "asset://primitives/primitive/label-pill?text=Onboard&accent=%23FF0000"
    )
    assert resolved.meta.ref.name == "label-pill"
    assert isinstance(resolved.payload, NativeShapePayload)


def test_payload_is_native_shape() -> None:
    payload = _resolve("metric-badge", (("value", "42"),))
    assert payload.primitive == "metric-badge"
    assert isinstance(payload.params, tuple)
