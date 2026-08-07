"""A2 Task 5 — deferred SVG theme materialization."""

import xml.etree.ElementTree as ET

import pytest

from offipy.assets import SvgPayload, SvgTemplatePayload
from offipy.assets.materialize import materialize_svg_template, resolve_asset_color
from offipy.exceptions import InvalidArgumentError


def _template(svg: str, slots):
    return SvgTemplatePayload(svg, "svg", (0, 0, 100, 100), slots)


# ---------------------------------------------------------------------------
# resolve_asset_color
# ---------------------------------------------------------------------------


class TestResolveAssetColor:
    def test_explicit_hex_bypasses_theme(self):
        assert resolve_asset_color("#2251ff", {}) == "#2251FF"

    def test_transparent_passthrough(self):
        assert resolve_asset_color("transparent", {}) == "transparent"

    def test_token_lookup_plain_key(self):
        assert resolve_asset_color("accent", {"accent": "#0052FF"}) == "#0052FF"

    def test_token_lookup_css_key(self):
        assert resolve_asset_color("accent", {"--accent": "#0052FF"}) == "#0052FF"

    def test_token_lookup_uppercases_hex(self):
        assert resolve_asset_color("accent", {"accent": "#0052ff"}) == "#0052FF"

    def test_bg_canonicalized_to_background(self):
        assert resolve_asset_color("bg", {"background": "#FFFFFF"}) == "#FFFFFF"

    def test_missing_token_fails(self):
        with pytest.raises(InvalidArgumentError):
            resolve_asset_color("accent", {})

    def test_theme_value_not_hex_fails(self):
        with pytest.raises(InvalidArgumentError):
            resolve_asset_color("accent", {"accent": "red"})
        # theme vars must not chain to another token
        with pytest.raises(InvalidArgumentError):
            resolve_asset_color("accent", {"accent": "ink"})

    def test_invalid_value_fails(self):
        with pytest.raises(InvalidArgumentError):
            resolve_asset_color("var(--accent)", {})
        with pytest.raises(InvalidArgumentError):
            resolve_asset_color("", {})

    def test_empty_theme_value_reports_not_defined(self):
        # deck 侧 measure 对未定义 CSS 变量 getPropertyValue 返回空串而非 None；
        # 空值应按「token 未定义」报错，不能落到「must be #RRGGBB」格式误报。
        with pytest.raises(InvalidArgumentError, match="not defined"):
            resolve_asset_color("accent", {"accent": ""})


# ---------------------------------------------------------------------------
# materialize_svg_template
# ---------------------------------------------------------------------------


class TestMaterializeSvgTemplate:
    def test_substitutes_theme_colors(self):
        tpl = _template(
            '<svg viewBox="0 0 100 100"><path fill="__A__"/><rect fill="__B__"/></svg>',
            (("__A__", "accent"), ("__B__", "ink")),
        )
        out = materialize_svg_template(tpl, {"accent": "#0052FF", "ink": "#111111"})
        assert isinstance(out, SvgPayload)
        assert out.render_mode == "svg"
        assert out.view_box == (0, 0, 100, 100)
        assert "__A__" not in out.svg
        assert "__B__" not in out.svg
        assert "#0052FF" in out.svg
        assert "#111111" in out.svg

    def test_explicit_hex_in_slot_bypasses_theme(self):
        tpl = _template('<svg><path fill="__A__"/></svg>', (("__A__", "#2251FF"),))
        out = materialize_svg_template(tpl, {})
        assert "#2251FF" in out.svg

    def test_transparent_slot(self):
        tpl = _template('<svg><path fill="__A__"/></svg>', (("__A__", "transparent"),))
        out = materialize_svg_template(tpl, {})
        assert 'fill="transparent"' in out.svg

    def test_light_and_dark_maps(self):
        tpl = _template('<svg><path fill="__A__"/></svg>', (("__A__", "accent"),))
        dark = materialize_svg_template(tpl, {"accent": "#0052FF"})
        light = materialize_svg_template(tpl, {"accent": "#66B3FF"})
        assert "#0052FF" in dark.svg
        assert "#66B3FF" in light.svg

    def test_missing_token_fails(self):
        tpl = _template('<svg><path fill="__A__"/></svg>', (("__A__", "accent"),))
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {})

    def test_declared_sentinel_absent_in_template_fails(self):
        tpl = _template('<svg><path fill="__A__"/></svg>', (("__A__", "accent"), ("__B__", "ink")))
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {"accent": "#0052FF", "ink": "#111111"})

    def test_undeclared_sentinel_in_template_fails(self):
        tpl = _template(
            '<svg><path fill="__A__"/><circle fill="__OFFIPY_X__"/></svg>', (("__A__", "accent"),)
        )
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {"accent": "#0052FF"})

    def test_malformed_svg_fails(self):
        tpl = _template('<svg><path fill="__A__"></svg>', (("__A__", "accent"),))
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {"accent": "#0052FF"})

    def test_output_deterministic(self):
        tpl = _template(
            '<svg><path fill="__A__"/><rect fill="__B__"/></svg>',
            (("__A__", "accent"), ("__B__", "ink")),
        )
        theme = {"accent": "#0052FF", "ink": "#111111"}
        a = materialize_svg_template(tpl, theme).svg
        b = materialize_svg_template(tpl, theme).svg
        assert a == b

    def test_materialized_svg_parses(self):
        tpl = _template('<svg><path fill="__A__"/></svg>', (("__A__", "accent"),))
        out = materialize_svg_template(tpl, {"accent": "#0052FF"})
        ET.fromstring(out.svg)
