"""A2 Task 1 — frozen asset core data model."""

import dataclasses

import pytest

from offipy.assets import (
    AssetKind,
    AssetMeta,
    AssetPlacement,
    AssetProviderMeta,
    AssetRect,
    AssetRef,
    AssetRenderContext,
    AssetRequest,
    NativeShapePayload,
    RasterPayload,
    ResolvedAsset,
    SvgPayload,
    SvgTemplatePayload,
)
from offipy.assets.model import canonical_params
from offipy.exceptions import InvalidArgumentError


# ---------------------------------------------------------------------------
# AssetRef
# ---------------------------------------------------------------------------

class TestAssetRef:
    def test_equality_hash(self):
        a = AssetRef("ph", "icon", "chart-line")
        b = AssetRef("ph", "icon", "chart-line")
        c = AssetRef("ph", "icon", "chart-bar")
        assert a == b
        assert hash(a) == hash(b)
        assert a != c

    def test_frozen(self):
        ref = AssetRef("ph", "icon", "chart-line")
        with pytest.raises(dataclasses.FrozenInstanceError):
            ref.name = "x"

    def test_provider_uppercase_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRef("Ph", "icon", "chart-line")

    def test_provider_underscore_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRef("my_provider", "icon", "chart-line")

    def test_kind_unknown_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRef("ph", "bogus", "chart-line")  # type: ignore[arg-type]

    def test_name_empty_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRef("ph", "icon", "")

    def test_name_path_segments_rejected(self):
        for bad in ("a/b", "a\\b", "..", "\x00"):
            with pytest.raises(InvalidArgumentError):
                AssetRef("ph", "icon", bad)

    def test_name_inner_space_allowed(self):
        AssetRef("ph", "icon", "a b")

    def test_name_leading_trailing_ws_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRef("ph", "icon", "  x")


# ---------------------------------------------------------------------------
# AssetRequest / canonical_params
# ---------------------------------------------------------------------------

class TestCanonicalParams:
    def test_sort_by_key(self):
        assert canonical_params([("b", "2"), ("a", "1")]) == (("a", "1"), ("b", "2"))

    def test_underscore_to_hyphen(self):
        assert canonical_params([("orb_count", "2")]) == (("orb-count", "2"),)

    def test_key_lowercased(self):
        assert canonical_params([("Seed", "4")]) == (("seed", "4"),)

    def test_key_stripped(self):
        assert canonical_params([("  seed ", "4")]) == (("seed", "4"),)

    def test_duplicate_canonical_key_fails(self):
        with pytest.raises(InvalidArgumentError):
            canonical_params([("seed", "1"), ("seed", "2")])
        with pytest.raises(InvalidArgumentError):
            canonical_params([("orb_count", "2"), ("orb-count", "3")])

    def test_empty_key_fails(self):
        with pytest.raises(InvalidArgumentError):
            canonical_params([("", "1")])
        with pytest.raises(InvalidArgumentError):
            canonical_params([(" ", "1")])

    def test_empty_value_allowed(self):
        assert canonical_params([("seed", "")]) == (("seed", ""),)

    def test_invalid_key_chars_rejected(self):
        with pytest.raises(InvalidArgumentError):
            canonical_params([("seed=1", "2")])
        with pytest.raises(InvalidArgumentError):
            canonical_params([("1abc", "2")])


class TestAssetRequest:
    def test_equality_independent_of_order(self):
        a = AssetRequest(AssetRef("procedural", "pattern", "topography"),
                         (("seed", "42"), ("foreground", "accent")))
        b = AssetRequest(AssetRef("procedural", "pattern", "topography"),
                         (("foreground", "accent"), ("seed", "42")))
        assert a == b
        assert hash(a) == hash(b)

    def test_direct_params_canonicalized(self):
        req = AssetRequest(AssetRef("procedural", "pattern", "topography"),
                           (("Seed", "4"), ("orb_count", "2")))
        assert req.params == (("orb-count", "2"), ("seed", "4"))

    def test_duplicate_key_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRequest(AssetRef("procedural", "pattern", "topography"),
                         (("seed", "1"), ("seed", "2")))

    def test_frozen(self):
        req = AssetRequest(AssetRef("ph", "icon", "chart-line"))
        with pytest.raises(dataclasses.FrozenInstanceError):
            req.params = ()


# ---------------------------------------------------------------------------
# Payloads
# ---------------------------------------------------------------------------

class TestSvgPayload:
    def test_render_mode_svg_ok(self):
        p = SvgPayload("<svg/>", "svg", (0, 0, 100, 100))
        assert p.render_mode == "svg"

    def test_render_mode_freeform_ok(self):
        p = SvgPayload("<svg/>", "freeform_svg", None)
        assert p.render_mode == "freeform_svg"

    def test_raster_mode_rejected(self):
        with pytest.raises(InvalidArgumentError):
            SvgPayload("<svg/>", "raster", None)  # type: ignore[arg-type]

    def test_native_shape_mode_rejected(self):
        with pytest.raises(InvalidArgumentError):
            SvgPayload("<svg/>", "native_shape", None)  # type: ignore[arg-type]

    def test_frozen(self):
        p = SvgPayload("<svg/>", "svg", None)
        with pytest.raises(dataclasses.FrozenInstanceError):
            p.svg = "x"


class TestSvgTemplatePayload:
    def test_render_mode_only_svg(self):
        with pytest.raises(InvalidArgumentError):
            SvgTemplatePayload("<svg/>", "freeform_svg", None, (("__A__", "accent"),))
        with pytest.raises(InvalidArgumentError):
            SvgTemplatePayload("<svg/>", "raster", None, (("__A__", "accent"),))

    def test_color_slots_unique_sentinels(self):
        with pytest.raises(InvalidArgumentError):
            SvgTemplatePayload("<svg/>", "svg", None,
                               (("__A__", "accent"), ("__A__", "ink")))

    def test_color_slots_empty_placeholder_rejected(self):
        with pytest.raises(InvalidArgumentError):
            SvgTemplatePayload("<svg/>", "svg", None, (("", "accent"),))


class TestRasterPayload:
    def test_preserves_fields(self):
        p = RasterPayload(b"\x89PNG", "image/png", 400, 240)
        assert p.data == b"\x89PNG"
        assert p.media_type == "image/png"
        assert p.pixel_width == 400
        assert p.pixel_height == 240

    def test_frozen(self):
        p = RasterPayload(b"x", "image/png", 1, 1)
        with pytest.raises(dataclasses.FrozenInstanceError):
            p.data = b"y"


class TestNativeShapePayload:
    def test_params_canonicalized(self):
        p = NativeShapePayload("metric-badge", (("Seed", "1"), ("width", "2")))
        assert p.params == (("seed", "1"), ("width", "2"))

    def test_duplicate_rejected(self):
        with pytest.raises(InvalidArgumentError):
            NativeShapePayload("metric-badge", (("seed", "1"), ("seed", "2")))


# ---------------------------------------------------------------------------
# ResolvedAsset
# ---------------------------------------------------------------------------

class TestResolvedAsset:
    def _resolved(self):
        ref = AssetRef("ph", "icon", "chart-line")
        req = AssetRequest(ref)
        return ResolvedAsset(
            request=req,
            meta=AssetMeta(ref=ref, title="Chart line", tags=("chart",)),
            provider_meta=AssetProviderMeta(
                provider_id="ph", license="ISC", source_url="https://example.com",
                source_commit="abc", attribution=None, redistributable=True,
                first_party=False,
            ),
            payload=SvgPayload("<svg/>", "freeform_svg", None),
        )

    def test_exactly_one_payload_variant(self):
        r = self._resolved()
        assert isinstance(r.payload, SvgPayload)
        assert not isinstance(r.payload, SvgTemplatePayload)
        assert not isinstance(r.payload, RasterPayload)
        assert not isinstance(r.payload, NativeShapePayload)

    def test_request_meta_provider_consistent(self):
        r = self._resolved()
        assert r.request.ref == r.meta.ref
        assert r.provider_meta.provider_id == r.request.ref.provider


# ---------------------------------------------------------------------------
# AssetRect / AssetRenderContext
# ---------------------------------------------------------------------------

class TestAssetRect:
    def test_unit_default_px(self):
        r = AssetRect(0, 0, 400, 240)
        assert r.unit == "px"

    def test_unit_non_px_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRect(0, 0, 400, 240, unit="pt")  # type: ignore[arg-type]

    def test_validate_render_positive(self):
        AssetRect(0, 0, 400, 240).validate_render()
        with pytest.raises(InvalidArgumentError):
            AssetRect(0, 0, 0, 240).validate_render()
        with pytest.raises(InvalidArgumentError):
            AssetRect(0, 0, 400, -1).validate_render()


class TestAssetRenderContext:
    def test_theme_vars_mapping(self):
        ctx = AssetRenderContext(
            slide_index=0,
            rect=AssetRect(0, 0, 100, 100),
            theme_name="mckinsey",
            theme_vars={"accent": "#0052FF"},
            placement="background",
        )
        assert ctx.theme_vars["accent"] == "#0052FF"
        assert ctx.placement == "background"

    def test_placement_rejects_unknown(self):
        with pytest.raises(InvalidArgumentError):
            AssetRenderContext(
                slide_index=0, rect=AssetRect(0, 0, 100, 100),
                theme_name=None, theme_vars={},
                placement="overlay",  # type: ignore[arg-type]
            )
