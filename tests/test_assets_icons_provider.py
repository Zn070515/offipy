"""A3 Task 1 — vendored Phosphor/Lucide icon providers."""

import json
from pathlib import Path

import pytest

from offipy.assets import (
    AssetRef,
    AssetRequest,
    SvgPayload,
    get_default_registry,
)
from offipy.assets.providers.icons import IconProvider, icon_render_mode
from offipy.exceptions import InvalidArgumentError

_REPO_ROOT = Path(__file__).resolve().parents[1]
_ICONS_DIR = _REPO_ROOT / "src" / "offipy" / "assets" / "icons"
_MANIFEST = json.loads((_ICONS_DIR / "manifest.json").read_text(encoding="utf-8"))
_MANIFEST_KEY = {"ph": "phosphor", "lu": "lucide"}


def _req(provider: str, name: str, params: tuple[tuple[str, str], ...] = ()) -> AssetRequest:
    return AssetRequest(AssetRef(provider, "icon", name), params)


# ---------------------------------------------------------------------------
# provider identity + internal freeform mode lookup
# ---------------------------------------------------------------------------


class TestProviderIdentity:
    def test_provider_ids_and_kinds(self):
        assert IconProvider("ph").provider_id == "ph"
        assert IconProvider("lu").provider_id == "lu"
        assert IconProvider("ph").kinds == frozenset({"icon"})
        assert IconProvider("lu").kinds == frozenset({"icon"})

    def test_unknown_provider_rejected(self):
        with pytest.raises(InvalidArgumentError):
            IconProvider("bogus")

    def test_icon_render_mode_lookup(self):
        assert icon_render_mode("ph") == "fill"
        assert icon_render_mode("lu") == "stroke"
        with pytest.raises(InvalidArgumentError):
            icon_render_mode("bogus")


class TestProviderMeta:
    @pytest.mark.parametrize("provider_id", ["ph", "lu"])
    def test_meta_maps_manifest_exactly(self, provider_id):
        meta = IconProvider(provider_id).provider_meta
        entry = _MANIFEST[_MANIFEST_KEY[provider_id]]
        assert meta.provider_id == provider_id
        assert meta.license == entry["license"]
        assert meta.source_url == entry["source"]
        assert meta.source_commit == entry["commit"]
        assert meta.attribution is None
        assert meta.redistributable is True
        assert meta.first_party is False


# ---------------------------------------------------------------------------
# deterministic search
# ---------------------------------------------------------------------------


class TestSearch:
    def test_phosphor_count_matches_manifest(self):
        ph = IconProvider("ph")
        metas = ph.search("", limit=2000)
        assert len(metas) == _MANIFEST["phosphor"]["count"]
        names = [m.ref.name for m in metas]
        assert names == sorted(names)

    def test_lucide_count_matches_manifest(self):
        lu = IconProvider("lu")
        metas = lu.search("", limit=2000)
        assert len(metas) == _MANIFEST["lucide"]["count"]
        names = [m.ref.name for m in metas]
        assert names == sorted(names)

    def test_search_is_deterministic(self):
        ph = IconProvider("ph")
        assert ph.search("") == ph.search("")

    def test_limit_applied(self):
        assert len(IconProvider("ph").search("", limit=5)) == 5

    def test_query_filters_by_substring(self):
        metas = IconProvider("lu").search("arrow", limit=100)
        assert metas
        assert all("arrow" in m.ref.name for m in metas)

    def test_kind_filter(self):
        ph = IconProvider("ph")
        assert ph.search("", kind="icon", limit=5)
        assert ph.search("", kind="pattern", limit=5) == []

    def test_title_and_tags_derived_deterministically(self):
        metas = IconProvider("ph").search("address", limit=10)
        assert metas
        for m in metas:
            assert m.title == m.ref.name
            assert m.tags == tuple(m.ref.name.split("-"))

    def test_ph_canonical_names_strip_fill_suffix(self):
        metas = IconProvider("ph").search("check", limit=20)
        assert metas
        assert all(not n.endswith("-fill") for n in (m.ref.name for m in metas))


# ---------------------------------------------------------------------------
# resolve
# ---------------------------------------------------------------------------


class TestResolve:
    def test_ph_check_returns_freeform_svg(self):
        prov = IconProvider("ph")
        req = _req("ph", "check")
        resolved = prov.resolve(req)
        assert resolved.request == req
        assert resolved.meta.ref == req.ref
        payload = resolved.payload
        assert isinstance(payload, SvgPayload)
        assert payload.render_mode == "freeform_svg"
        assert payload.view_box == (0.0, 0.0, 256.0, 256.0)
        assert "<svg" in payload.svg

    def test_lu_check_returns_svg(self):
        prov = IconProvider("lu")
        resolved = prov.resolve(_req("lu", "check"))
        assert isinstance(resolved.payload, SvgPayload)
        assert resolved.payload.render_mode == "freeform_svg"
        assert resolved.payload.view_box == (0.0, 0.0, 24.0, 24.0)

    def test_view_box_parsed_from_source_not_prefix(self):
        assert IconProvider("ph").resolve(_req("ph", "check")).payload.view_box == (
            0.0,
            0.0,
            256.0,
            256.0,
        )
        assert IconProvider("lu").resolve(_req("lu", "check")).payload.view_box == (
            0.0,
            0.0,
            24.0,
            24.0,
        )

    def test_ph_fill_suffix_alias_same_svg(self):
        prov = IconProvider("ph")
        a = prov.resolve(_req("ph", "check"))
        b = prov.resolve(_req("ph", "check-fill"))
        assert a.payload.svg == b.payload.svg

    def test_meta_title_and_tags(self):
        resolved = IconProvider("ph").resolve(_req("ph", "check"))
        assert resolved.meta.title == "check"
        assert resolved.meta.tags == ("check",)

    def test_missing_icon_rejected(self):
        with pytest.raises(InvalidArgumentError):
            IconProvider("ph").resolve(_req("ph", "no-such-icon"))

    def test_any_params_rejected(self):
        with pytest.raises(InvalidArgumentError):
            IconProvider("ph").resolve(_req("ph", "check", (("seed", "1"),)))
        with pytest.raises(InvalidArgumentError):
            IconProvider("lu").resolve(_req("lu", "check", (("color", "accent"),)))

    def test_wrong_kind_rejected(self):
        req = AssetRequest(AssetRef("ph", "pattern", "check"))
        with pytest.raises(InvalidArgumentError):
            IconProvider("ph").resolve(req)

    def test_traversal_name_rejected_at_ref(self):
        with pytest.raises(InvalidArgumentError):
            AssetRef("ph", "icon", "..%2Fetc")

    def test_resolve_reads_vendored_file(self):
        # the provider must resolve purely from vendored files, no network
        payload = IconProvider("ph").resolve(_req("ph", "airplane")).payload
        expected = (_ICONS_DIR / "phosphor" / "airplane-fill.svg").read_text(encoding="utf-8")
        assert payload.svg == expected


# ---------------------------------------------------------------------------
# default registry integration
# ---------------------------------------------------------------------------


class TestDefaultRegistryIcons:
    def test_registers_ph_then_lu_in_order(self):
        reg = get_default_registry()
        assert reg.provider("ph").provider_id == "ph"
        assert reg.provider("lu").provider_id == "lu"

    def test_default_registry_resolves_icon_uri(self):
        resolved = get_default_registry().resolve("asset://ph/icon/check")
        assert isinstance(resolved.payload, SvgPayload)
        assert resolved.payload.render_mode == "freeform_svg"

    def test_search_counts_match_manifest_sum(self):
        metas = get_default_registry().search("", limit=10000)
        assert len(metas) == _MANIFEST["phosphor"]["count"] + _MANIFEST["lucide"]["count"]
