"""A4 Tasks 2+10 — procedural provider contracts and resolve -> SvgTemplatePayload.

Full boundary table, strict type conversion, unknown-param rejection, color
syntax via the A2 helper, deterministic search order, provider metadata, and
resolve wiring into the pattern builders with conditional color slots.
"""

from __future__ import annotations

import pytest

from offipy.assets.license import LicensePolicy
from offipy.assets.materialize import materialize_svg_template
from offipy.assets.model import AssetRef, AssetRequest, SvgTemplatePayload
from offipy.assets.patterns._common import BG, FG
from offipy.assets.providers.procedural import (
    _PATTERN_ORDER,
    ProceduralProvider,
    _coerce_params,
    _parse_float,
    _parse_int,
)
from offipy.assets.registry import get_default_registry
from offipy.exceptions import InvalidArgumentError

_INT32_MIN = -(1 << 31)
_INT32_MAX = (1 << 31) - 1

# (pattern, param, low, high) — int bounds are ints, float bounds are floats.
_BOUNDARY_CASES: tuple[tuple[str, str, int | float, int | float], ...] = (
    ("wave", "density", 0.0, 1.0),
    ("wave", "thickness", 0.0, 1.0),
    ("blob", "complexity", 2, 8),
    ("dot-grid", "spacing", 0.5, 2.0),
    ("dot-grid", "radius", 0.0, 0.5),
    ("square-grid", "spacing", 0.5, 2.0),
    ("square-grid", "thickness", 0.0, 1.0),
    ("rings", "count", 1, 12),
    ("rings", "thickness", 0.0, 1.0),
    ("topography", "density", 0.0, 1.0),
    ("topography", "lines", 3, 24),
    ("circuit", "nodes", 4, 60),
    ("circuit", "density", 0.0, 1.0),
    ("gradient-orb", "orb-count", 1, 6),
    ("gradient-orb", "blur", 0.0, 1.0),
)

_DEFAULT_CASES: tuple[tuple[str, dict[str, int | float]], ...] = (
    ("wave", {"density": 0.5, "thickness": 0.5}),
    ("blob", {"complexity": 5}),
    ("dot-grid", {"spacing": 1.0, "radius": 0.15}),
    ("square-grid", {"spacing": 1.0, "thickness": 0.5}),
    ("rings", {"count": 5, "thickness": 0.5}),
    ("topography", {"density": 0.5, "lines": 10}),
    ("circuit", {"nodes": 20, "density": 0.5}),
    ("gradient-orb", {"orb-count": 3, "blur": 0.5}),
)


# -- shared params ---------------------------------------------------------


@pytest.mark.parametrize(("pattern", "expected"), _DEFAULT_CASES)
def test_defaults_applied(pattern: str, expected: dict[str, int | float]) -> None:
    got = _coerce_params(pattern, ())
    assert got["seed"] == 0
    assert got["foreground"] == "accent"
    assert got["background"] == "transparent"
    for key, value in expected.items():
        if isinstance(value, int):
            assert got[key] == value
        else:
            assert got[key] == pytest.approx(value)


@pytest.mark.parametrize(("pattern", "param", "low", "high"), _BOUNDARY_CASES)
def test_boundary_min_max_accepted(pattern: str, param: str, low: float, high: float) -> None:
    for edge in (low, high):
        got = _coerce_params(pattern, ((param, str(edge)),))
        if isinstance(edge, int):
            assert got[param] == edge
        else:
            assert got[param] == pytest.approx(float(edge))


@pytest.mark.parametrize(("pattern", "param", "low", "high"), _BOUNDARY_CASES)
def test_boundary_outside_rejected(pattern: str, param: str, low: float, high: float) -> None:
    bad = [low - 1, high + 1] if isinstance(low, int) else [low - 0.001, high + 0.001]
    for value in bad:
        with pytest.raises(InvalidArgumentError):
            _coerce_params(pattern, ((param, str(value)),))


def test_int_grammar_accepts_decimal_forms() -> None:
    assert _parse_int("-5", "x", -10, 10) == -5
    assert _parse_int("0", "x", -10, 10) == 0
    assert _parse_int("007", "x", -10, 10) == 7
    assert _parse_int("-0", "x", -10, 10) == 0


def test_float_grammar_accepts_decimal_forms() -> None:
    for value, expected in [
        ("0", 0.0),
        (".5", 0.5),
        ("5.", 5.0),
        ("-0.5", -0.5),
        ("1.5", 1.5),
        ("-0", 0.0),
    ]:
        assert _parse_float(value, "x", -10, 10) == expected


@pytest.mark.parametrize(
    "value",
    ["1.0", "1e3", " 5", "5 ", "", "abc", "0x5", "+5", "5..0", "0.5"],
)
def test_int_grammar_rejects_non_integers(value: str) -> None:
    with pytest.raises(InvalidArgumentError):
        _coerce_params("blob", (("complexity", value),))


@pytest.mark.parametrize(
    "value",
    ["1e3", "nan", "inf", "-inf", " 0.5", "0.5 ", "", "abc", "0x0.1p0", "+0.5", "--0.5"],
)
def test_float_grammar_rejects_non_decimals(value: str) -> None:
    with pytest.raises(InvalidArgumentError):
        _coerce_params("wave", (("density", value),))


def test_seed_bounds_accepted() -> None:
    got = _coerce_params("wave", (("seed", str(_INT32_MIN)),))
    assert got["seed"] == _INT32_MIN
    got = _coerce_params("wave", (("seed", str(_INT32_MAX)),))
    assert got["seed"] == _INT32_MAX


@pytest.mark.parametrize("value", ["2147483648", "-2147483649", "1e3", "0x10", ""])
def test_seed_bounds_rejected(value: str) -> None:
    with pytest.raises(InvalidArgumentError):
        _coerce_params("wave", (("seed", value),))


def test_unknown_pattern_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="unknown procedural pattern"):
        _coerce_params("nope", ())


def test_unknown_param_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="unknown param"):
        _coerce_params("wave", (("bogus", "1"),))


def test_unknown_param_for_other_pattern_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="unknown param"):
        _coerce_params("wave", (("orb-count", "3"),))


def test_unknown_param_message_lists_allowed_keys() -> None:
    with pytest.raises(InvalidArgumentError, match="allowed"):
        _coerce_params("wave", (("bogus", "1"),))


def test_duplicate_param_rejected() -> None:
    with pytest.raises(InvalidArgumentError, match="duplicate"):
        _coerce_params("wave", (("seed", "1"), ("seed", "2")))


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("accent", "accent"),
        ("accent-soft", "accent-soft"),
        ("ink", "ink"),
        ("muted", "muted"),
        ("surface", "surface"),
        ("background", "background"),
        ("bg", "background"),
        ("transparent", "transparent"),
        ("#ff00aa", "#FF00AA"),
        ("#F0A", "#F0A"),
    ],
)
def test_color_syntax_accepted(value: str, expected: str) -> None:
    got = _coerce_params("wave", (("foreground", value),))
    assert got["foreground"] == expected


@pytest.mark.parametrize("value", ["#GGHHII", "red", "var(--accent)", "", " rgb(1,2,3)"])
def test_color_syntax_rejected(value: str) -> None:
    with pytest.raises(InvalidArgumentError):
        _coerce_params("wave", (("foreground", value),))


def test_background_canonicalized() -> None:
    got = _coerce_params("wave", (("background", "#abcdef"),))
    assert got["background"] == "#ABCDEF"


# -- search ----------------------------------------------------------------


def test_search_returns_all_patterns_in_order() -> None:
    names = [m.ref.name for m in ProceduralProvider().search("")]
    assert names == list(_PATTERN_ORDER)


def test_search_limit() -> None:
    assert len(ProceduralProvider().search("", limit=3)) == 3


def test_search_substring_case_insensitive() -> None:
    provider = ProceduralProvider()
    assert [m.ref.name for m in provider.search("grid")] == ["dot-grid", "square-grid"]
    assert [m.ref.name for m in provider.search("GRID")] == ["dot-grid", "square-grid"]


def test_search_matches_title_and_tags() -> None:
    provider = ProceduralProvider()
    assert [m.ref.name for m in provider.search("concentric")] == ["rings"]
    assert [m.ref.name for m in provider.search("tech")] == ["circuit"]
    assert [m.ref.name for m in provider.search("dot")] == ["dot-grid"]


def test_search_no_match() -> None:
    assert ProceduralProvider().search("zzz") == []


def test_search_kind_filter() -> None:
    provider = ProceduralProvider()
    assert provider.search("", kind="icon") == []
    assert len(provider.search("", kind="pattern")) == 8


# -- metadata / resolve ----------------------------------------------------


def test_provider_metadata() -> None:
    provider = ProceduralProvider()
    assert provider.provider_id == "procedural"
    assert provider.kinds == frozenset({"pattern"})
    meta = provider.provider_meta
    assert meta.license == "MIT"
    assert meta.first_party is True
    assert meta.redistributable is True
    assert meta.source_url == "https://github.com/Zn070515/offipy"
    assert meta.source_commit is None
    assert meta.attribution is None


def test_provider_meta_passes_license_policy() -> None:
    LicensePolicy().validate_provider_meta(ProceduralProvider().provider_meta)


def test_resolve_rejects_non_pattern_kind() -> None:
    req = AssetRequest(AssetRef("procedural", "icon", "wave"))
    with pytest.raises(InvalidArgumentError):
        ProceduralProvider().resolve(req)


def test_resolve_rejects_unknown_pattern() -> None:
    req = AssetRequest(AssetRef("procedural", "pattern", "nope"))
    with pytest.raises(InvalidArgumentError):
        ProceduralProvider().resolve(req)


# -- resolve -> SvgTemplatePayload (Task 10) --------------------------------


def test_resolve_default_returns_svg_template() -> None:
    req = AssetRequest(AssetRef("procedural", "pattern", "wave"))
    resolved = ProceduralProvider().resolve(req)
    assert isinstance(resolved.payload, SvgTemplatePayload)
    assert resolved.payload.render_mode == "svg"
    assert resolved.payload.view_box == (0.0, 0.0, 1000.0, 1000.0)
    assert resolved.request == req
    assert resolved.meta.ref == req.ref


@pytest.mark.parametrize("pattern", list(_PATTERN_ORDER))
def test_resolve_every_pattern_default(pattern: str) -> None:
    req = AssetRequest(AssetRef("procedural", "pattern", pattern))
    resolved = ProceduralProvider().resolve(req)
    payload = resolved.payload
    assert isinstance(payload, SvgTemplatePayload)
    assert dict(payload.color_slots)[FG] == "accent"


def test_color_slots_exactly_cover_template_sentinels() -> None:
    for pattern in _PATTERN_ORDER:
        resolved = ProceduralProvider().resolve(
            AssetRequest(AssetRef("procedural", "pattern", pattern))
        )
        payload = resolved.payload
        assert isinstance(payload, SvgTemplatePayload)
        slots = dict(payload.color_slots)
        for sentinel in (FG, BG):
            assert (sentinel in payload.template) == (sentinel in slots)


def test_dot_grid_radius_zero_omits_fg_slot() -> None:
    req = AssetRequest(AssetRef("procedural", "pattern", "dot-grid"), (("radius", "0"),))
    payload = ProceduralProvider().resolve(req).payload
    assert isinstance(payload, SvgTemplatePayload)
    assert FG not in payload.template
    assert FG not in dict(payload.color_slots)


def test_resolve_params_flow_into_template() -> None:
    req = AssetRequest(
        AssetRef("procedural", "pattern", "gradient-orb"),
        (("orb-count", "6"), ("blur", "1")),
    )
    payload = ProceduralProvider().resolve(req).payload
    assert isinstance(payload, SvgTemplatePayload)
    assert payload.template.count("<circle") == 6
    assert "<radialGradient" in payload.template
    assert 'id="orb-5"' in payload.template


def test_resolved_template_materializes() -> None:
    req = AssetRequest(AssetRef("procedural", "pattern", "wave"), (("seed", "7"),))
    resolved = ProceduralProvider().resolve(req)
    assert isinstance(resolved.payload, SvgTemplatePayload)
    svg = materialize_svg_template(resolved.payload, {"accent": "#112233"}).svg
    assert "__OFFIPY_ASSET" not in svg
    assert "#112233" in svg


def test_explicit_default_same_template_different_request() -> None:
    provider = ProceduralProvider()
    plain = provider.resolve(AssetRequest(AssetRef("procedural", "pattern", "rings")))
    explicit = provider.resolve(
        AssetRequest(AssetRef("procedural", "pattern", "rings"), (("seed", "0"),))
    )
    assert isinstance(plain.payload, SvgTemplatePayload)
    assert isinstance(explicit.payload, SvgTemplatePayload)
    assert plain.payload.template == explicit.payload.template
    assert plain.request != explicit.request


def test_resolve_meta_title_and_tags() -> None:
    resolved = ProceduralProvider().resolve(
        AssetRequest(AssetRef("procedural", "pattern", "topography"))
    )
    assert resolved.meta.title == "Topography Contours"
    assert "contour" in resolved.meta.tags


def test_resolve_keeps_explicit_color_values() -> None:
    req = AssetRequest(
        AssetRef("procedural", "pattern", "rings"),
        (("foreground", "#ff00aa"), ("background", "#123456")),
    )
    payload = ProceduralProvider().resolve(req).payload
    assert isinstance(payload, SvgTemplatePayload)
    slots = dict(payload.color_slots)
    assert slots[FG] == "#FF00AA"
    assert slots[BG] == "#123456"


def test_default_registry_resolves_procedural_uri() -> None:
    resolved = get_default_registry().resolve("asset://procedural/pattern/wave?seed=1")
    assert isinstance(resolved.payload, SvgTemplatePayload)
    assert resolved.provider_meta.first_party is True
    assert resolved.provider_meta.license == "MIT"
