"""offipy.assets — frozen asset core data model.

A2 contract (rev1.2 §3): identity, request, discriminated payloads, metadata,
provenance, render context. Pure stdlib; must not import python-pptx/Pillow/Playwright.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal, Protocol

from offipy.exceptions import InvalidArgumentError

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping

AssetKind = Literal[
    "icon",
    "illustration",
    "pattern",
    "primitive",
    "map",
    "flag",
    "logo",
    "photo",
]
AssetRenderMode = Literal["freeform_svg", "svg", "raster", "native_shape"]
AssetPlacement = Literal["replace", "background", "decorative"]

_ASSET_KINDS: frozenset[str] = frozenset(
    {
        "icon",
        "illustration",
        "pattern",
        "primitive",
        "map",
        "flag",
        "logo",
        "photo",
    }
)
_PLACEMENTS: frozenset[str] = frozenset({"replace", "background", "decorative"})

_PROVIDER_RE = re.compile(r"[a-z][a-z0-9-]*\Z")
_PARAM_KEY_RE = re.compile(r"[a-z][a-z0-9-]*\Z")


def canonical_params(items: Iterable[tuple[str, str]]) -> tuple[tuple[str, str], ...]:
    """Canonicalize an iterable of (key, value) pairs.

    Rules (rev1.2 §3.2): strip + lowercase key, `_` canonicalized to `-`, valid
    key regex `[a-z][a-z0-9-]*`, duplicate canonical key fails, sort by key
    ascending, values stay strings (empty allowed). Keys stay in the given order
    when equal — but equal keys are rejected, so the result is fully deterministic.
    """
    canon: dict[str, str] = {}
    for key, value in items:
        k = key.strip().lower().replace("_", "-")
        if not k:
            raise InvalidArgumentError("empty asset param key")
        if not _PARAM_KEY_RE.match(k):
            raise InvalidArgumentError(f"invalid asset param key {key!r}")
        if k in canon:
            raise InvalidArgumentError(f"duplicate asset param key {k!r}")
        canon[k] = value
    return tuple(sorted(canon.items()))


def _check_kind(kind: str) -> None:
    if kind not in _ASSET_KINDS:
        raise InvalidArgumentError(f"unknown asset kind {kind!r}")


def _check_placement(placement: str) -> None:
    if placement not in _PLACEMENTS:
        raise InvalidArgumentError(f"invalid asset placement {placement!r}")


@dataclass(frozen=True)
class AssetRef:
    provider: str
    kind: AssetKind
    name: str

    def __post_init__(self) -> None:
        if not _PROVIDER_RE.match(self.provider):
            raise InvalidArgumentError(f"invalid provider id {self.provider!r}")
        _check_kind(self.kind)
        name = self.name
        if not name:
            raise InvalidArgumentError("asset name must be non-empty")
        if name != name.strip():
            raise InvalidArgumentError("asset name must not have surrounding whitespace")
        if "/" in name or "\\" in name or ".." in name or "\x00" in name:
            raise InvalidArgumentError(f"invalid asset name {name!r}")


@dataclass(frozen=True)
class AssetRequest:
    ref: AssetRef
    params: tuple[tuple[str, str], ...] = ()

    def __post_init__(self) -> None:
        if self.params:
            object.__setattr__(self, "params", canonical_params(self.params))


@dataclass(frozen=True)
class AssetMeta:
    ref: AssetRef
    title: str
    tags: tuple[str, ...]
    editable: bool = False
    trademark: bool = False


@dataclass(frozen=True)
class AssetProviderMeta:
    provider_id: str
    license: str
    source_url: str | None
    source_commit: str | None
    attribution: str | None
    redistributable: bool
    first_party: bool = False


@dataclass(frozen=True)
class SvgPayload:
    svg: str
    render_mode: Literal["freeform_svg", "svg"]
    view_box: tuple[float, float, float, float] | None

    def __post_init__(self) -> None:
        if self.render_mode not in ("freeform_svg", "svg"):
            raise InvalidArgumentError(
                f"SvgPayload render_mode must be freeform_svg or svg, got {self.render_mode!r}"
            )


@dataclass(frozen=True)
class SvgTemplatePayload:
    template: str
    render_mode: Literal["svg"]
    view_box: tuple[float, float, float, float] | None
    color_slots: tuple[tuple[str, str], ...]

    def __post_init__(self) -> None:
        if self.render_mode != "svg":
            raise InvalidArgumentError(
                f"SvgTemplatePayload render_mode must be svg, got {self.render_mode!r}"
            )
        seen: set[str] = set()
        for placeholder, _value in self.color_slots:
            if not placeholder:
                raise InvalidArgumentError("color_slots placeholder must be non-empty")
            if placeholder in seen:
                raise InvalidArgumentError(f"duplicate color_slots placeholder {placeholder!r}")
            seen.add(placeholder)


@dataclass(frozen=True)
class RasterPayload:
    data: bytes
    media_type: str
    pixel_width: int
    pixel_height: int


@dataclass(frozen=True)
class NativeShapePayload:
    primitive: str
    params: tuple[tuple[str, str], ...]

    def __post_init__(self) -> None:
        if not self.primitive:
            raise InvalidArgumentError("primitive name must be non-empty")
        if self.params:
            object.__setattr__(self, "params", canonical_params(self.params))


AssetPayload = SvgPayload | SvgTemplatePayload | RasterPayload | NativeShapePayload


@dataclass(frozen=True)
class ResolvedAsset:
    request: AssetRequest
    meta: AssetMeta
    provider_meta: AssetProviderMeta
    payload: AssetPayload


@dataclass(frozen=True)
class AssetRect:
    x: float
    y: float
    width: float
    height: float
    unit: Literal["px"] = "px"

    def __post_init__(self) -> None:
        if self.unit != "px":
            raise InvalidArgumentError(f"AssetRect unit must be px, got {self.unit!r}")

    def validate_render(self) -> None:
        if self.width <= 0 or self.height <= 0:
            raise InvalidArgumentError(
                f"AssetRect width/height must be positive for render, got "
                f"{self.width}x{self.height}"
            )


@dataclass(frozen=True)
class AssetRenderContext:
    slide_index: int
    rect: AssetRect
    theme_name: str | None
    theme_vars: Mapping[str, str]
    placement: AssetPlacement

    def __post_init__(self) -> None:
        _check_placement(self.placement)
        # keep the frozen context independent of the caller's mutable mapping
        object.__setattr__(self, "theme_vars", dict(self.theme_vars))


@dataclass(frozen=True)
class RenderedAssetElements:
    """Renderer output; internal-facing. Elements are OOXML XML elements."""

    elements: tuple[object, ...]


class AssetProvider(Protocol):
    provider_id: str
    kinds: frozenset[AssetKind]
    provider_meta: AssetProviderMeta

    def search(
        self, query: str, *, kind: AssetKind | None = None, limit: int = 20
    ) -> list[AssetMeta]: ...
    def resolve(self, request: AssetRequest) -> ResolvedAsset: ...
