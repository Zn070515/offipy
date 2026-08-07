"""A2 Task 2 — canonical asset URI parsing/formatting + color value syntax."""

import pytest

from offipy.assets import AssetRef, AssetRequest
from offipy.assets.color import validate_color_value
from offipy.assets.uri import format_asset_uri, parse_asset_uri
from offipy.exceptions import InvalidArgumentError


class TestParseValid:
    def test_icon_uri(self):
        req = parse_asset_uri("asset://ph/icon/chart-line")
        assert req == AssetRequest(AssetRef("ph", "icon", "chart-line"))
        assert req.params == ()

    def test_procedural_query(self):
        req = parse_asset_uri("asset://procedural/pattern/topography?seed=42&foreground=accent")
        assert req.ref == AssetRef("procedural", "pattern", "topography")
        assert req.params == (("foreground", "accent"), ("seed", "42"))

    def test_query_order_independent(self):
        a = parse_asset_uri("asset://procedural/pattern/topography?seed=42&foreground=accent")
        b = parse_asset_uri("asset://procedural/pattern/topography?foreground=accent&seed=42")
        assert a == b
        assert hash(a) == hash(b)

    def test_canonical_format_sorted(self):
        uri = format_asset_uri(
            parse_asset_uri("asset://procedural/pattern/topography?seed=42&foreground=accent")
        )
        assert uri == "asset://procedural/pattern/topography?foreground=accent&seed=42"

    def test_hex_color_encoded(self):
        req = parse_asset_uri("asset://procedural/pattern/topography?foreground=%232251FF")
        assert req.params == (("foreground", "#2251FF"),)
        assert format_asset_uri(req) == "asset://procedural/pattern/topography?foreground=%232251FF"

    def test_roundtrip(self):
        for uri in (
            "asset://ph/icon/chart-line",
            "asset://procedural/pattern/topography?seed=42&foreground=accent",
            "asset://procedural/pattern/topography?foreground=%232251FF",
        ):
            req = parse_asset_uri(uri)
            assert parse_asset_uri(format_asset_uri(req)) == req

    def test_parse_accepts_request(self):
        # registry.resolve may be given a request directly; parser contract only
        # applies to strings — this just guards the string branch stays simple.
        assert parse_asset_uri("asset://lu/icon/database") == AssetRequest(
            AssetRef("lu", "icon", "database")
        )


class TestParseInvalid:
    def test_wrong_scheme(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("http://ph/icon/chart-line")

    def test_missing_segments(self):
        for uri in ("asset://ph", "asset://ph/icon", "asset://"):
            with pytest.raises(InvalidArgumentError):
                parse_asset_uri(uri)

    def test_fourth_segment(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph/icon/chart-line/extra")

    def test_blank_segments(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset:///icon/chart-line")
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph//chart-line")

    def test_unknown_kind(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph/bogus/chart-line")

    def test_fragment_rejected(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph/icon/chart-line#frag")

    def test_naked_hash_hex_becomes_fragment(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph/icon/chart-line?foreground=#2251FF")

    def test_duplicate_query_key(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://procedural/pattern/topography?seed=1&seed=2")

    def test_duplicate_canonical_underscore_hyphen(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://procedural/pattern/topography?orb_count=2&orb-count=3")

    def test_var_value_rejected(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://procedural/pattern/topography?foreground=var(--accent)")

    def test_malformed_percent_escape(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph/icon/chart-line?seed=%ZZ")

    def test_path_traversal_after_decode(self):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri("asset://ph/icon/..%2Fetc")

    def test_errors_are_invalid_argument(self):
        for uri in ("asset://ph/icon/x?a=%2", "asset://ph/icon/", "asset://PH/icon/x"):
            with pytest.raises(InvalidArgumentError):
                parse_asset_uri(uri)


class TestValidateColorValue:
    def test_tokens(self):
        assert validate_color_value("accent") == "accent"
        assert validate_color_value("accent-soft") == "accent-soft"
        assert validate_color_value("ink") == "ink"
        assert validate_color_value("muted") == "muted"
        assert validate_color_value("surface") == "surface"
        assert validate_color_value("background") == "background"
        assert validate_color_value("transparent") == "transparent"

    def test_bg_canonicalized(self):
        assert validate_color_value("bg") == "background"

    def test_hex_normalized_uppercase(self):
        assert validate_color_value("#2251ff") == "#2251FF"
        assert validate_color_value("#ABC") == "#ABC"

    def test_hex_3_digit_accepted(self):
        assert validate_color_value("#abc") == "#ABC"

    def test_invalid(self):
        for bad in (
            "var(--accent)",
            "",
            "notacolor",
            "#12345",
            "#12345G",
            "#GGGGGG",
            "##",
            "accent ",
        ):
            with pytest.raises(InvalidArgumentError):
                validate_color_value(bad)
