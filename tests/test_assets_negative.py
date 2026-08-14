"""A2 Task 7 — core fuzz/negative contract suite.

Every user-input error must surface as `InvalidArgumentError`; provider
contract violations must name the provider id in the message.
"""

import pytest

from offipy.assets import (
    AssetMeta,
    AssetProviderMeta,
    AssetRect,
    AssetRenderContext,
    ResolvedAsset,
    SvgPayload,
    SvgTemplatePayload,
    materialize_svg_template,
    parse_asset_uri,
)
from offipy.assets.license import LicensePolicy
from offipy.assets.model import canonical_params
from offipy.assets.registry import AssetRegistry
from offipy.exceptions import InvalidArgumentError

_POLICY = LicensePolicy()


# ---------------------------------------------------------------------------
# malformed asset URIs (≥30 table)
# ---------------------------------------------------------------------------

_MALFORMED_URIS = [
    "",
    "ph/icon/chart",
    "asset:/ph/icon/chart",
    "asset://",
    "asset://ph",
    "asset://ph/icon",
    "asset://ph/icon/chart/extra",
    "asset:///icon/chart",
    "asset://ph//chart",
    "asset://ph/icon/",
    "asset://ph/icon/chart/",
    "asset://PH/icon/chart",
    "asset://ph/ICON/chart",
    "asset://ph/icon/chart#frag",
    "asset://ph/icon/chart#",
    "asset://ph:8080/icon/chart",
    "asset://ph/icon/chart?seed=1&seed=2",
    "asset://ph/icon/chart?seed=1&Seed=2",
    "asset://ph/icon/chart?orb_count=1&orb-count=2",
    "asset://ph/icon/chart?=1",
    "asset://ph/icon/chart?1abc=1",
    "asset://ph/icon/chart?seed..x=1",
    "asset://ph/icon/chart?seed=%ZZ",
    "asset://ph/icon/chart?seed=%2",
    "asset://ph/icon/chart?seed=var(--accent)",
    "asset://ph/icon/chart?seed=var(--x)",
    "asset://ph/icon/chart?..=1",
    "asset://ph/icon/..%2Fetc",
    "asset://ph/icon/%2e%2e",
    "asset://ph/icon/a%2Fb",
    "asset://ph/icon/a%5Cb",
    "asset://ph/icon/a%00b",
    "asset://ph/icon/ chart",
    "asset://ph/icon/chart ",
    "asset://ph/icon/chart?seed=%FF",
]


class TestMalformedAssetUris:
    @pytest.mark.parametrize("uri", _MALFORMED_URIS)
    def test_rejects(self, uri):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri(uri)


# ---------------------------------------------------------------------------
# duplicate query / canonical-key collisions
# ---------------------------------------------------------------------------

_COLLIDING_URIS = [
    "asset://procedural/pattern/topo?seed=1&seed=2",
    "asset://procedural/pattern/topo?seed=1&Seed=2",
    "asset://procedural/pattern/topo?orb_count=1&orb-count=2",
]


class TestQueryCollisions:
    @pytest.mark.parametrize("uri", _COLLIDING_URIS)
    def test_parse_rejects(self, uri):
        with pytest.raises(InvalidArgumentError):
            parse_asset_uri(uri)

    def test_canonical_params_rejects_duplicates(self):
        with pytest.raises(InvalidArgumentError):
            canonical_params([("seed", "1"), ("seed", "2")])
        with pytest.raises(InvalidArgumentError):
            canonical_params([("orb_count", "1"), ("orb-count", "2")])


# ---------------------------------------------------------------------------
# bad provider implementations (contract violations name the provider)
# ---------------------------------------------------------------------------


def _provider_meta(provider_id: str) -> AssetProviderMeta:
    return AssetProviderMeta(
        provider_id=provider_id,
        license="ISC",
        source_url=None,
        source_commit=None,
        attribution=None,
        redistributable=True,
    )


class _BadProvider:
    def __init__(self, provider_id="ph", kinds=("icon",), *, bad_payload=False):
        self.provider_id = provider_id
        self.kinds = frozenset(kinds)
        self.provider_meta = _provider_meta(provider_id)
        self._bad_payload = bad_payload

    def search(self, query, *, kind=None, limit=20):
        return []

    def resolve(self, request):
        if self._bad_payload:
            return ResolvedAsset(
                request=request,
                meta=AssetMeta(request.ref, request.ref.name, ()),
                provider_meta=self.provider_meta,
                payload="nope",  # type: ignore[arg-type]
            )
        return ResolvedAsset(
            request=request,
            meta=AssetMeta(request.ref, request.ref.name, ()),
            provider_meta=self.provider_meta,
            payload=SvgPayload("<svg/>", "svg", None),
        )


class TestBadProviderContract:
    def test_register_unknown_kind_names_provider(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError, match="ph"):
            reg.register(_BadProvider(kinds=("bogus",)))

    def test_resolve_kind_mismatch_names_provider(self):
        reg = AssetRegistry()
        reg.register(_BadProvider(kinds=("icon",)))
        with pytest.raises(InvalidArgumentError, match="ph"):
            reg.resolve("asset://ph/pattern/topo")

    def test_resolve_bad_payload_names_provider(self):
        reg = AssetRegistry()
        reg.register(_BadProvider(bad_payload=True))
        with pytest.raises(InvalidArgumentError, match="ph"):
            reg.resolve("asset://ph/icon/a")


# ---------------------------------------------------------------------------
# bad license metadata
# ---------------------------------------------------------------------------


def _meta_with(**kw) -> AssetProviderMeta:
    defaults = {
        "provider_id": "ph",
        "license": "ISC",
        "source_url": None,
        "source_commit": None,
        "attribution": None,
        "redistributable": True,
    }
    defaults.update(kw)
    return AssetProviderMeta(**defaults)


_BAD_PROVIDER_META = [
    {"license": "GPL-3.0"},
    {"license": "Apache-2.0"},
    {"license": ""},
    {"license": "CC-BY-4.0"},
    {"license": "CC-BY-4.0", "source_url": "https://x"},
    {"license": "CC-BY-4.0", "attribution": "A"},
]


class TestBadLicenseMetadata:
    @pytest.mark.parametrize("kw", _BAD_PROVIDER_META)
    def test_provider_meta_rejected(self, kw):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_provider_meta(_meta_with(**kw))


_BAD_MANIFESTS = [
    {},
    {"license": "MIT", "source": "x"},
    {"license": "MIT", "source": "x", "source_commit": "c", "count": 1, "redistributable": False},
    {"license": "MIT", "source": "", "source_commit": "c", "count": 1},
    {"license": "MIT", "source": "x", "source_commit": "", "count": 1},
    {"license": "MIT", "source": "x", "source_commit": "c", "count": -1},
    {"license": "MIT", "source": "x", "source_commit": "c", "count": "1"},
    {"license": "Apache-2.0", "source": "x", "source_commit": "c", "count": 1},
    {"license": "CC-BY-4.0", "source": "x", "source_commit": "c", "count": 1},
]


class TestBadLicenseManifests:
    @pytest.mark.parametrize("data", _BAD_MANIFESTS)
    def test_manifest_rejected(self, data):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(data)


# ---------------------------------------------------------------------------
# malformed SVG template sentinels
# ---------------------------------------------------------------------------


class TestBadSvgTemplate:
    def _tpl(self, svg: str, slots):
        return SvgTemplatePayload(svg, "svg", None, slots)

    def test_undeclared_sentinel(self):
        tpl = self._tpl('<svg><circle fill="__OFFIPY_X__"/></svg>', (("__A__", "accent"),))
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {"accent": "#0052FF"})

    def test_declared_but_absent(self):
        tpl = self._tpl("<svg/>", (("__A__", "accent"),))
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {"accent": "#0052FF"})

    def test_malformed_svg(self):
        tpl = self._tpl("<svg><path></svg>", (("__A__", "accent"),))
        with pytest.raises(InvalidArgumentError):
            materialize_svg_template(tpl, {"accent": "#0052FF"})


# ---------------------------------------------------------------------------
# invalid rect units / dimensions at the validation boundary
# ---------------------------------------------------------------------------


class TestRectValidationBoundary:
    def test_negative_and_zero_dimensions(self):
        for x, y, w, h in [(0, 0, -1, 10), (0, 0, 10, -1), (0, 0, 0, 10)]:
            rect = AssetRect(x, y, w, h)
            with pytest.raises(InvalidArgumentError):
                rect.validate_render()

    def test_non_px_unit_rejected(self):
        with pytest.raises(InvalidArgumentError):
            AssetRect(0, 0, 10, 10, unit="pt")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# mutable mapping accidentally passed into frozen context
# ---------------------------------------------------------------------------


class TestFrozenContextMappingDiscipline:
    def test_theme_vars_copied_on_construct(self):
        vars_map = {"accent": "#0052FF"}
        ctx = AssetRenderContext(
            slide_index=0,
            rect=AssetRect(0, 0, 100, 100),
            theme_name=None,
            theme_vars=vars_map,
            placement="background",
        )
        vars_map["accent"] = "#000000"
        assert ctx.theme_vars["accent"] == "#0052FF"
