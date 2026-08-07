"""offipy.assets — deterministic provider registry.

Contract (rev1.2 §3.8): providers register with validated identity/kinds;
search is deterministic (registration order, provider-local order, dedup by
AssetRef with first occurrence winning) and applies a global limit; resolve
parses a URI or accepts an AssetRequest and runs an integrity gate on the
provider's result so one bad provider cannot corrupt assets.json/provenance.
"""

from __future__ import annotations

from offipy.assets.model import (
    _ASSET_KINDS,
    _PROVIDER_RE,
    AssetKind,
    AssetMeta,
    AssetProvider,
    AssetRequest,
    NativeShapePayload,
    RasterPayload,
    ResolvedAsset,
    SvgPayload,
    SvgTemplatePayload,
)
from offipy.assets.uri import parse_asset_uri
from offipy.exceptions import InvalidArgumentError

_PAYLOAD_VARIANTS = (
    SvgPayload,
    SvgTemplatePayload,
    RasterPayload,
    NativeShapePayload,
)


class AssetRegistry:
    """Holds providers in registration order and serves deterministic lookups."""

    def __init__(self) -> None:
        self._providers: dict[str, AssetProvider] = {}

    def register(self, provider: AssetProvider) -> None:
        self._validate_provider(provider)
        provider_id = provider.provider_id
        if provider_id in self._providers:
            raise InvalidArgumentError(f"duplicate asset provider {provider_id!r}")
        self._providers[provider_id] = provider

    def _validate_provider(self, provider: AssetProvider) -> None:
        provider_id = provider.provider_id
        if not _PROVIDER_RE.match(provider_id):
            raise InvalidArgumentError(f"invalid asset provider id {provider_id!r}")
        kinds = provider.kinds
        if not kinds:
            raise InvalidArgumentError(f"asset provider {provider_id!r} declares no kinds")
        for kind in kinds:
            if kind not in _ASSET_KINDS:
                raise InvalidArgumentError(
                    f"asset provider {provider_id!r} declares unknown kind {kind!r}"
                )
        if provider.provider_meta.provider_id != provider_id:
            raise InvalidArgumentError(
                f"asset provider {provider_id!r} provider_meta.provider_id "
                f"mismatch {provider.provider_meta.provider_id!r}"
            )

    def provider(self, provider_id: str) -> AssetProvider:
        try:
            return self._providers[provider_id]
        except KeyError:
            raise InvalidArgumentError(f"unknown asset provider {provider_id!r}") from None

    def search(
        self, query: str, *, kind: AssetKind | None = None, limit: int = 20
    ) -> list[AssetMeta]:
        if limit <= 0:
            raise InvalidArgumentError(f"search limit must be positive, got {limit}")
        results: list[AssetMeta] = []
        seen: set[tuple[str, str, str]] = set()
        for provider in self._providers.values():
            if kind is not None and kind not in provider.kinds:
                continue
            for meta in provider.search(query, kind=kind, limit=limit):
                ref = meta.ref
                if ref.provider != provider.provider_id:
                    raise InvalidArgumentError(
                        f"asset provider {provider.provider_id!r} returned meta for "
                        f"provider {ref.provider!r}"
                    )
                if kind is not None and ref.kind != kind:
                    raise InvalidArgumentError(
                        f"asset provider {provider.provider_id!r} returned kind "
                        f"{ref.kind!r} outside filter {kind!r}"
                    )
                if ref.kind not in provider.kinds:
                    raise InvalidArgumentError(
                        f"asset provider {provider.provider_id!r} returned kind "
                        f"{ref.kind!r} not in its declared kinds"
                    )
                key = (ref.provider, ref.kind, ref.name)
                if key in seen:
                    continue
                seen.add(key)
                results.append(meta)
                if len(results) >= limit:
                    return results
        return results

    def resolve(self, uri_or_request: str | AssetRequest) -> ResolvedAsset:
        if isinstance(uri_or_request, str):
            request = parse_asset_uri(uri_or_request)
        elif isinstance(uri_or_request, AssetRequest):
            request = uri_or_request
        else:
            raise InvalidArgumentError("resolve expects an asset URI string or AssetRequest")
        provider = self.provider(request.ref.provider)
        resolved = provider.resolve(request)
        self._validate_resolved(provider, request, resolved)
        return resolved

    def _validate_resolved(
        self, provider: AssetProvider, request: AssetRequest, resolved: ResolvedAsset
    ) -> None:
        if not isinstance(resolved, ResolvedAsset):
            raise InvalidArgumentError(
                f"asset provider {provider.provider_id!r} did not return a ResolvedAsset"
            )
        if resolved.request != request:
            raise InvalidArgumentError(
                f"asset provider {provider.provider_id!r} resolved a different request than asked"
            )
        if resolved.meta.ref != request.ref:
            raise InvalidArgumentError(
                f"asset provider {provider.provider_id!r} returned meta.ref "
                f"mismatch {resolved.meta.ref}"
            )
        if resolved.provider_meta.provider_id != request.ref.provider:
            raise InvalidArgumentError(
                f"asset provider {provider.provider_id!r} returned provider_meta "
                f"mismatch {resolved.provider_meta.provider_id!r}"
            )
        if request.ref.kind not in provider.kinds:
            raise InvalidArgumentError(
                f"asset provider {provider.provider_id!r} does not support kind "
                f"{request.ref.kind!r}"
            )
        if not isinstance(resolved.payload, _PAYLOAD_VARIANTS):
            raise InvalidArgumentError(
                f"asset provider {provider.provider_id!r} returned invalid payload "
                f"type {type(resolved.payload).__name__!r}"
            )


_default_registry: AssetRegistry | None = None


def _build_default_registry() -> AssetRegistry:
    reg = AssetRegistry()
    # vendored icon providers register in deterministic order; provider
    # construction does not scan the asset directories.
    from offipy.assets.providers.icons import IconProvider

    reg.register(IconProvider("ph"))
    reg.register(IconProvider("lu"))

    from offipy.assets.providers.procedural import ProceduralProvider

    reg.register(ProceduralProvider())

    from offipy.assets.providers.primitives import PrimitivesProvider

    reg.register(PrimitivesProvider())
    return reg


def get_default_registry() -> AssetRegistry:
    """Return the process-wide default registry, constructing it lazily."""
    global _default_registry
    if _default_registry is None:
        _default_registry = _build_default_registry()
    return _default_registry


def reset_default_registry_for_tests() -> None:
    """Drop the cached default registry (test/private use only)."""
    global _default_registry
    _default_registry = None
