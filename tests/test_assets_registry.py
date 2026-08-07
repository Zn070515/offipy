"""A2 Task 3 — deterministic provider registry."""

import pytest

from offipy.assets import (
    AssetMeta,
    AssetProviderMeta,
    AssetRef,
    AssetRequest,
    ResolvedAsset,
    SvgPayload,
)
from offipy.assets.registry import (
    AssetRegistry,
    get_default_registry,
    reset_default_registry_for_tests,
)
from offipy.exceptions import InvalidArgumentError


def _meta(ref: AssetRef, title: str | None = None) -> AssetMeta:
    return AssetMeta(ref=ref, title=title or ref.name, tags=())


def _provider_meta(provider_id: str) -> AssetProviderMeta:
    return AssetProviderMeta(
        provider_id=provider_id, license="ISC", source_url=None,
        source_commit=None, attribution=None, redistributable=True,
    )


class FakeProvider:
    def __init__(self, provider_id="ph", kinds=("icon",), results=None, *,
                 resolve_override=None):
        self.provider_id = provider_id
        self.kinds = frozenset(kinds)
        self.provider_meta = _provider_meta(provider_id)
        self._results = list(results or [])
        self._resolve_override = resolve_override
        self.search_calls: list[tuple[str, object, int]] = []

    def search(self, query, *, kind=None, limit=20):
        self.search_calls.append((query, kind, limit))
        out = []
        for m in self._results:
            if kind is not None and m.ref.kind != kind:
                continue
            if query and query not in m.title:
                continue
            out.append(m)
        return out[:limit]

    def resolve(self, request):
        if self._resolve_override is not None:
            return self._resolve_override(request)
        return ResolvedAsset(
            request=request,
            meta=_meta(request.ref),
            provider_meta=self.provider_meta,
            payload=SvgPayload("<svg/>", "svg", None),
        )


# ---------------------------------------------------------------------------
# register / provider
# ---------------------------------------------------------------------------

class TestRegister:
    def test_register_and_provider_roundtrip(self):
        reg = AssetRegistry()
        p = FakeProvider("ph")
        reg.register(p)
        assert reg.provider("ph") is p

    def test_duplicate_provider_id_rejected(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph"))
        with pytest.raises(InvalidArgumentError):
            reg.register(FakeProvider("ph"))

    def test_blank_provider_id_rejected(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError):
            reg.register(FakeProvider(""))

    def test_noncanonical_provider_id_rejected(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError):
            reg.register(FakeProvider("MyProvider"))

    def test_empty_kinds_rejected(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError):
            reg.register(FakeProvider("ph", kinds=()))

    def test_unknown_kind_in_kinds_rejected(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError):
            reg.register(FakeProvider("ph", kinds=("bogus",)))

    def test_provider_meta_mismatch_rejected(self):
        reg = AssetRegistry()
        p = FakeProvider("ph")
        p.provider_meta = _provider_meta("other")
        with pytest.raises(InvalidArgumentError):
            reg.register(p)

    def test_provider_not_found(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError):
            reg.provider("ph")


# ---------------------------------------------------------------------------
# search
# ---------------------------------------------------------------------------

class TestSearch:
    def test_empty_query_allowed(self):
        reg = AssetRegistry()
        p = FakeProvider("ph", results=[_meta(AssetRef("ph", "icon", "a"))])
        reg.register(p)
        assert len(reg.search("")) == 1
        assert p.search_calls == [("", None, 20)]

    def test_registration_order_preserved(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph", results=[_meta(AssetRef("ph", "icon", "a"))]))
        reg.register(FakeProvider("lu", results=[_meta(AssetRef("lu", "icon", "b"))]))
        assert [m.ref.provider for m in reg.search("")] == ["ph", "lu"]

    def test_provider_local_order_preserved(self):
        reg = AssetRegistry()
        p = FakeProvider("ph", results=[
            _meta(AssetRef("ph", "icon", "a")),
            _meta(AssetRef("ph", "icon", "b")),
        ])
        reg.register(p)
        assert [m.ref.name for m in reg.search("")] == ["a", "b"]

    def test_dedup_same_ref_first_wins(self):
        reg = AssetRegistry()
        first = _meta(AssetRef("ph", "icon", "a"))
        dup = _meta(AssetRef("ph", "icon", "a"), title="duplicate")
        reg.register(FakeProvider("ph", results=[
            first, dup, _meta(AssetRef("ph", "icon", "b"))]))
        metas = reg.search("")
        assert len(metas) == 2
        assert metas[0] is first  # first occurrence wins, dup dropped

    def test_global_limit_not_per_provider(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph", results=[
            _meta(AssetRef("ph", "icon", f"a{i}")) for i in range(5)]))
        reg.register(FakeProvider("lu", results=[
            _meta(AssetRef("lu", "icon", f"b{i}")) for i in range(5)]))
        metas = reg.search("", limit=7)
        assert len(metas) == 7
        assert sum(1 for m in metas if m.ref.provider == "ph") == 5
        assert sum(1 for m in metas if m.ref.provider == "lu") == 2

    def test_limit_zero_or_negative_fails(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph"))
        with pytest.raises(InvalidArgumentError):
            reg.search("", limit=0)
        with pytest.raises(InvalidArgumentError):
            reg.search("", limit=-1)

    def test_kind_filter_skips_incompatible_providers(self):
        reg = AssetRegistry()
        p_icon = FakeProvider("ph", kinds=("icon",),
                              results=[_meta(AssetRef("ph", "icon", "a"))])
        p_pat = FakeProvider("lu", kinds=("pattern",),
                             results=[_meta(AssetRef("lu", "pattern", "b"))])
        reg.register(p_icon)
        reg.register(p_pat)
        metas = reg.search("", kind="pattern")
        assert [m.ref.provider for m in metas] == ["lu"]
        assert p_icon.search_calls == []

    def test_oversized_provider_truncated_safely(self):
        class Unbounded(FakeProvider):
            def search(self, query, *, kind=None, limit=20):
                return list(self._results)
        reg = AssetRegistry()
        reg.register(Unbounded("ph", results=[
            _meta(AssetRef("ph", "icon", f"a{i}")) for i in range(5)]))
        assert len(reg.search("", limit=2)) == 2

    def test_wrong_provider_ref_in_meta_rejected(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph", results=[_meta(AssetRef("other", "icon", "a"))]))
        with pytest.raises(InvalidArgumentError):
            reg.search("")


# ---------------------------------------------------------------------------
# resolve
# ---------------------------------------------------------------------------

class TestResolve:
    def test_resolve_by_uri_string(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph"))
        resolved = reg.resolve("asset://ph/icon/chart-line")
        assert isinstance(resolved, ResolvedAsset)
        assert resolved.request == AssetRequest(AssetRef("ph", "icon", "chart-line"))

    def test_resolve_by_request(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph"))
        req = AssetRequest(AssetRef("ph", "icon", "chart-line"))
        assert reg.resolve(req).request == req

    def test_resolve_unknown_provider(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError):
            reg.resolve("asset://ph/icon/chart-line")

    def test_resolve_invalid_uri(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph"))
        with pytest.raises(InvalidArgumentError):
            reg.resolve("http://ph/icon/chart-line")

    def test_resolve_kind_not_in_provider_kinds(self):
        reg = AssetRegistry()
        reg.register(FakeProvider("ph", kinds=("icon",)))
        with pytest.raises(InvalidArgumentError):
            reg.resolve("asset://ph/pattern/topo")

    def test_resolve_wrong_request_rejected(self):
        reg = AssetRegistry()
        def bad(request):
            return ResolvedAsset(
                request=AssetRequest(AssetRef("ph", "icon", "OTHER")),
                meta=_meta(request.ref),
                provider_meta=_provider_meta("ph"),
                payload=SvgPayload("<svg/>", "svg", None),
            )
        reg.register(FakeProvider("ph", resolve_override=bad))
        with pytest.raises(InvalidArgumentError):
            reg.resolve("asset://ph/icon/chart-line")

    def test_resolve_wrong_meta_ref_rejected(self):
        reg = AssetRegistry()
        def bad(request):
            return ResolvedAsset(
                request=request,
                meta=_meta(AssetRef("ph", "icon", "OTHER")),
                provider_meta=_provider_meta("ph"),
                payload=SvgPayload("<svg/>", "svg", None),
            )
        reg.register(FakeProvider("ph", resolve_override=bad))
        with pytest.raises(InvalidArgumentError):
            reg.resolve("asset://ph/icon/chart-line")

    def test_resolve_wrong_provider_meta_id_rejected(self):
        reg = AssetRegistry()
        def bad(request):
            return ResolvedAsset(
                request=request,
                meta=_meta(request.ref),
                provider_meta=_provider_meta("other"),
                payload=SvgPayload("<svg/>", "svg", None),
            )
        reg.register(FakeProvider("ph", resolve_override=bad))
        with pytest.raises(InvalidArgumentError):
            reg.resolve("asset://ph/icon/chart-line")

    def test_resolve_invalid_payload_rejected(self):
        reg = AssetRegistry()
        def bad(request):
            return ResolvedAsset(
                request=request,
                meta=_meta(request.ref),
                provider_meta=_provider_meta("ph"),
                payload="not a payload",  # type: ignore[assignment]
            )
        reg.register(FakeProvider("ph", resolve_override=bad))
        with pytest.raises(InvalidArgumentError):
            reg.resolve("asset://ph/icon/chart-line")


# ---------------------------------------------------------------------------
# default registry lifecycle
# ---------------------------------------------------------------------------

class TestDefaultRegistry:
    def test_get_default_registry_lazy_singleton(self):
        a = get_default_registry()
        b = get_default_registry()
        assert a is b
        assert isinstance(a, AssetRegistry)

    def test_reset_clears_singleton(self):
        reg = get_default_registry()
        reset_default_registry_for_tests()
        c = get_default_registry()
        assert c is not reg
