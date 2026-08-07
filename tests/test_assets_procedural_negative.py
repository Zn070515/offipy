"""A4 Task 16 — procedural provider negative / security hardening.

Query injection, sentinel/XML-smuggling color values, numeric extremes, unknown
pattern/param/kind/provider, malformed hex — all rejected before any geometry is
built. A full resolve → search → materialize cycle is verified to be fully
offline (no socket ever opens).
"""

import socket

import pytest

from offipy.assets.materialize import materialize_svg_template
from offipy.assets.providers.procedural import _PATTERN_ORDER
from offipy.assets.registry import get_default_registry
from offipy.assets.uri import parse_asset_uri
from offipy.exceptions import InvalidArgumentError


def _resolve(uri):
    return get_default_registry().resolve(uri)


class TestQueryInjection:
    def test_duplicate_query_key_rejected(self):
        with pytest.raises(InvalidArgumentError, match="duplicate"):
            parse_asset_uri("asset://procedural/pattern/wave?seed=1&seed=2")

    def test_percent_encoded_ampersand_value_rejected(self):
        # %26 是值里的字面 &，不是分隔符 → resolve 时该值不是合法 int
        with pytest.raises(InvalidArgumentError):
            _resolve("asset://procedural/pattern/wave?seed=1%26seed%3D2")

    def test_css_var_reference_rejected(self):
        with pytest.raises(InvalidArgumentError, match="var"):
            parse_asset_uri("asset://procedural/pattern/wave?foreground=var(%23accent)")

    def test_trailing_newline_in_int_rejected(self):
        with pytest.raises(InvalidArgumentError):
            _resolve("asset://procedural/pattern/wave?seed=1%0A")

    def test_query_ordering_is_canonical(self):
        a = parse_asset_uri("asset://procedural/pattern/wave?thickness=0.2&seed=3&density=0.5")
        b = parse_asset_uri("asset://procedural/pattern/wave?seed=3&density=0.5&thickness=0.2")
        assert a.params == b.params
        assert [k for k, _ in a.params] == ["density", "seed", "thickness"]  # 排序稳定


class TestColorSmuggling:
    def test_sentinel_color_injection_rejected(self):
        with pytest.raises(InvalidArgumentError):
            _resolve("asset://procedural/pattern/wave?foreground=__OFFIPY_ASSET_FG__")

    def test_xml_special_chars_in_color_rejected(self):
        for bad in ("%3Crect%3E", "a%26b", "%22quote%22", "%3Cscript%3E"):
            with pytest.raises(InvalidArgumentError):
                _resolve(f"asset://procedural/pattern/wave?foreground={bad}")

    def test_malformed_hex_rejected(self):
        for bad in ("%23GGHHII", "%2312345", "%231234567", "zzzz"):
            with pytest.raises(InvalidArgumentError):
                _resolve(f"asset://procedural/pattern/wave?foreground={bad}")


class TestNumericExtremes:
    def test_int_overflow_rejected(self):
        for value in ("2147483648", "-2147483649"):
            with pytest.raises(InvalidArgumentError):
                _resolve(f"asset://procedural/pattern/blob?seed={value}")

    def test_massive_numeric_string_rejected(self):
        with pytest.raises(InvalidArgumentError):
            _resolve(f"asset://procedural/pattern/wave?density={'9' * 1000}")

    def test_nan_inf_rejected(self):
        for value in ("nan", "inf", "-inf", "1e999"):
            with pytest.raises(InvalidArgumentError):
                _resolve(f"asset://procedural/pattern/wave?density={value}")

    def test_bounds_still_enforced_at_resolve(self):
        with pytest.raises(InvalidArgumentError):
            _resolve("asset://procedural/pattern/wave?density=1.001")


class TestUnknowns:
    def test_unknown_pattern_rejected(self):
        with pytest.raises(InvalidArgumentError, match="unknown procedural pattern"):
            _resolve("asset://procedural/pattern/bogus")

    def test_unknown_param_rejected(self):
        with pytest.raises(InvalidArgumentError, match="unknown param"):
            _resolve("asset://procedural/pattern/wave?bogus=1")

    def test_wrong_kind_rejected(self):
        with pytest.raises(InvalidArgumentError):
            _resolve("asset://procedural/icon/wave")

    def test_unknown_provider_rejected(self):
        with pytest.raises(InvalidArgumentError, match="unknown asset provider"):
            _resolve("asset://nope/pattern/wave")


class TestOffline:
    def test_resolve_search_materialize_never_opens_socket(self, monkeypatch):
        def _deny(*a, **k):
            raise AssertionError("network access attempted")

        for name in ("socket", "create_connection", "getaddrinfo"):
            monkeypatch.setattr(socket, name, _deny)

        registry = get_default_registry()
        for pattern in _PATTERN_ORDER:
            resolved = registry.resolve(f"asset://procedural/pattern/{pattern}?seed=3")
            svg = materialize_svg_template(resolved.payload, {"accent": "#112233"}).svg
            assert "__OFFIPY_ASSET" not in svg and "#112233" in svg
        assert len(registry.search("", kind="pattern", limit=2000)) == 8
