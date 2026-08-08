"""A3 Task 2 — HTML asset declaration parser / canonicalizer."""

import pytest

from offipy.assets.declarations import preprocess_asset_declarations
from offipy.assets.model import AssetRef
from offipy.exceptions import InvalidArgumentError


def _decls(html: str):
    return preprocess_asset_declarations(html)


_HTML = """<!doctype html>
<html><body>
<section data-pptx-slide>
  <h1>标题页</h1>
  <div data-asset="asset://ph/icon/check"></div>
  <svg data-icon="ph:chart-line"></svg>
</section>
<section data-pptx-slide>
  <div data-asset="asset://ph/icon/check"></div>
</section>
</body></html>
"""


class TestBasicExtraction:
    def test_declarations_extracted_in_order(self):
        _, decls = _decls(_HTML)
        assert len(decls) == 3
        d0, d1, d2 = decls
        assert (d0.slide_index, d0.declaration_id) == (1, "asset-s01-001")
        assert (d1.slide_index, d1.declaration_id) == (1, "asset-s01-002")
        assert (d2.slide_index, d2.declaration_id) == (2, "asset-s02-001")
        assert d0.request.ref == AssetRef("ph", "icon", "check")
        assert d1.request.ref == AssetRef("ph", "icon", "chart-line")
        assert d2.request.ref == AssetRef("ph", "icon", "check")
        assert d0.placement == "replace"
        assert d0.html_tag == "div"
        assert d1.html_tag == "svg"

    def test_legacy_icon_gets_canonical_attrs(self):
        rewritten, _ = _decls(_HTML)
        assert 'data-asset="asset://ph/icon/chart-line"' in rewritten
        assert 'data-offipy-asset-id="asset-s01-002"' in rewritten
        assert 'data-asset-placement="replace"' in rewritten

    def test_identical_icons_get_distinct_ids(self):
        _, decls = _decls(_HTML)
        ids = [d.declaration_id for d in decls]
        assert len(set(ids)) == len(ids)

    def test_rewrite_is_deterministic(self):
        a, _ = _decls(_HTML)
        b, _ = _decls(_HTML)
        assert a == b

    def test_rewrite_preserves_surrounding_text_byte_for_byte(self):
        rewritten, _ = _decls(_HTML)
        assert "标题页" in rewritten
        assert "<h1>标题页</h1>" in rewritten
        assert rewritten.count("asset-s01-001") == 1
        assert rewritten.count('data-asset="asset://ph/icon/check"') == 2


class TestParams:
    def test_uri_and_attr_params_merge_sorted(self):
        html = (
            "<section data-pptx-slide><div data-asset="
            '"asset://procedural/pattern/topo?seed=42" '
            'data-asset-param-foreground="accent"></div></section>'
        )
        _, [d] = _decls(html)
        assert d.request.params == (("foreground", "accent"), ("seed", "42"))

    def test_param_key_canonicalized(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://procedural/pattern/topo" '
            'data-asset-param-orb_count="3"></div></section>'
        )
        _, [d] = _decls(html)
        assert d.request.params == (("orb-count", "3"),)

    def test_unrelated_data_attrs_ignored(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://ph/icon/check" '
            'data-role="hero" data-layout="center" data-chart="line"></div></section>'
        )
        _, [d] = _decls(html)
        assert d.request.params == ()

    def test_uri_attr_param_collision_fails(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://procedural/pattern/topo?seed=1" '
            'data-asset-param-seed="2"></div></section>'
        )
        with pytest.raises(InvalidArgumentError):
            _decls(html)

    def test_empty_param_key_fails(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://ph/icon/check" '
            'data-asset-param-="1"></div></section>'
        )
        with pytest.raises(InvalidArgumentError):
            _decls(html)

    def test_bare_data_asset_param_without_suffix_ignored(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://ph/icon/check" '
            'data-asset-param="1"></div></section>'
        )
        _, [d] = _decls(html)
        assert d.request.params == ()


class TestPlacement:
    def test_default_replace(self):
        html = '<section data-pptx-slide><div data-asset="asset://ph/icon/check"></div></section>'
        _, [d] = _decls(html)
        assert d.placement == "replace"
        rewritten, _ = _decls(html)
        assert 'data-asset-placement="replace"' in rewritten

    @pytest.mark.parametrize("val", ["background", "decorative"])
    def test_explicit_placement(self, val):
        html = (
            f'<section data-pptx-slide><div data-asset="asset://ph/icon/check" '
            f'data-asset-placement="{val}"></div></section>'
        )
        _, [d] = _decls(html)
        assert d.placement == val

    def test_unknown_placement_fails_with_context(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://ph/icon/check" '
            'data-asset-placement="bogus"></div></section>'
        )
        with pytest.raises(InvalidArgumentError, match="slide"):
            _decls(html)


class TestPrimitiveSugar:
    def test_data_primitive_maps_to_primitives_provider(self):
        html = '<section data-pptx-slide><div data-primitive="quote-mark"></div></section>'
        _, [d] = _decls(html)
        assert d.request.ref == AssetRef("primitives", "primitive", "quote-mark")
        assert d.placement == "replace"
        rewritten, _ = _decls(html)
        assert 'data-asset="asset://primitives/primitive/quote-mark"' in rewritten

    def test_empty_primitive_fails(self):
        html = '<section data-pptx-slide><div data-primitive=""></div></section>'
        with pytest.raises(InvalidArgumentError):
            _decls(html)


class TestConflicts:
    @pytest.mark.parametrize(
        "conflict",
        [
            'data-asset="asset://ph/icon/check" data-icon="ph:check"',
            'data-asset="asset://ph/icon/check" data-primitive="quote-mark"',
            'data-icon="ph:check" data-primitive="quote-mark"',
        ],
    )
    def test_conflicting_declarations_fail(self, conflict):
        html = f"<section data-pptx-slide><div {conflict}></div></section>"
        with pytest.raises(InvalidArgumentError):
            _decls(html)

    def test_source_with_internal_id_fails(self):
        html = (
            '<section data-pptx-slide><div data-asset="asset://ph/icon/check" '
            'data-offipy-asset-id="asset-s01-001"></div></section>'
        )
        with pytest.raises(InvalidArgumentError):
            _decls(html)

    def test_declaration_outside_slide_fails(self):
        html = '<div data-asset="asset://ph/icon/check"></div><section data-pptx-slide></section>'
        with pytest.raises(InvalidArgumentError):
            _decls(html)

    def test_declaration_between_slides_fails(self):
        html = (
            "<section data-pptx-slide></section>"
            '<div data-asset="asset://ph/icon/check"></div>'
            "<section data-pptx-slide></section>"
        )
        with pytest.raises(InvalidArgumentError):
            _decls(html)

    def test_empty_data_asset_fails(self):
        html = '<section data-pptx-slide><div data-asset=""></div></section>'
        with pytest.raises(InvalidArgumentError):
            _decls(html)


class TestLegacyIconUriEncoding:
    def test_plain_name_roundtrips(self):
        rewritten, decls = _decls(
            '<section data-pptx-slide><svg data-icon="ph:check"></svg></section>'
        )
        assert 'data-asset="asset://ph/icon/check"' in rewritten
        assert decls[0].request.ref.name == "check"

    def test_name_needing_encoding_stays_canonical(self):
        # 名称含空格等必须百分号编码的字符：注入副本仍是规范 URI，不被截断
        rewritten, _ = _decls(
            '<section data-pptx-slide><svg data-icon="ph:arrow left"></svg></section>'
        )
        assert 'data-asset="asset://ph/icon/arrow%20left"' in rewritten

    def test_special_chars_do_not_become_uri_structure(self):
        # 若原样拼接，'?'/'#' 会改变 URI 结构（query/fragment）导致 parse 报
        # fragment 错误；编码后全部留在 name 数据段，交给下游名校验拒绝。
        rewritten, _ = _decls('<section data-pptx-slide><svg data-icon="ph:a?b#c"></svg></section>')
        assert 'data-asset="asset://ph/icon/a%3Fb%23c"' in rewritten

    def test_unknown_legacy_icon_set_fails(self):
        html = '<section data-pptx-slide><svg data-icon="tabler:x"></svg></section>'
        with pytest.raises(InvalidArgumentError):
            _decls(html)


class TestLegacyLayoutPreserved:
    def test_legacy_svg_attributes_unchanged(self):
        html = (
            "<section data-pptx-slide>"
            '<svg class="icon" data-icon="ph:check" viewBox="0 0 256 256" '
            'width="48" height="48"></svg>'
            "</section>"
        )
        rewritten, _ = _decls(html)
        assert 'class="icon"' in rewritten
        assert 'viewBox="0 0 256 256"' in rewritten
        assert 'width="48"' in rewritten
        assert 'height="48"' in rewritten
        assert 'data-icon="ph:check"' in rewritten
        assert 'data-asset="asset://ph/icon/check"' in rewritten
        assert 'data-offipy-asset-id="asset-s01-001"' in rewritten


class TestAttributeSyntax:
    def test_single_quoted_attrs(self):
        html = "<section data-pptx-slide><div data-asset='asset://ph/icon/check'></div></section>"
        rewritten, [d] = _decls(html)
        assert d.request.ref == AssetRef("ph", "icon", "check")
        assert 'data-offipy-asset-id="asset-s01-001"' in rewritten

    def test_self_closing_declaration(self):
        html = '<section data-pptx-slide><div data-asset="asset://ph/icon/check"/></section>'
        rewritten, [d] = _decls(html)
        assert d.request.ref == AssetRef("ph", "icon", "check")
        assert 'data-offipy-asset-id="asset-s01-001"' in rewritten


class TestSlideIndexing:
    def test_slide_index_is_section_ordinal(self):
        html = (
            "<header>noise</header>"
            '<section data-pptx-slide><div data-asset="asset://ph/icon/a"></div></section>'
            "<div>between</div>"
            '<section data-pptx-slide><div data-asset="asset://ph/icon/b"></div></section>'
        )
        _, decls = _decls(html)
        assert [d.slide_index for d in decls] == [1, 2]

    def test_same_slide_ordinals_increment(self):
        html = (
            "<section data-pptx-slide>"
            '<div data-asset="asset://ph/icon/a"></div>'
            '<div data-asset="asset://ph/icon/b"></div>'
            '<div data-asset="asset://ph/icon/c"></div>'
            "</section>"
        )
        _, decls = _decls(html)
        assert [d.declaration_id for d in decls] == [
            "asset-s01-001",
            "asset-s01-002",
            "asset-s01-003",
        ]


class TestNoDeclarations:
    def test_no_declarations_returns_identity(self):
        html = "<section data-pptx-slide><p>hi</p></section>"
        rewritten, decls = _decls(html)
        assert decls == []
        assert rewritten == html
