"""offipy.assets.providers.icons — vendored Phosphor/Lucide icon providers.

A3 Task 1: wrap the two vendored SVG sets in the AssetProvider protocol without
touching the SVG/freeform math. Pure stdlib; icons are read on demand from the
vendored directory and the provider never opens the network.

The fill/stroke render-mode distinction is freeform-engine metadata, not part of
the public Asset model; it is exposed here as `icon_render_mode`.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Literal

from offipy.assets._xml import parse_svg
from offipy.assets.model import (
    AssetKind,
    AssetMeta,
    AssetProviderMeta,
    AssetRef,
    AssetRequest,
    ResolvedAsset,
    SvgPayload,
)
from offipy.exceptions import InvalidArgumentError

_ICONS_DIR = Path(__file__).resolve().parent.parent.parent / "assets" / "icons"

_SET_DIRS = {"ph": "phosphor", "lu": "lucide"}
_SET_MODE: dict[str, Literal["fill", "stroke"]] = {"ph": "fill", "lu": "stroke"}
_NAME_RE = re.compile(r"[a-z0-9][a-z0-9-]*\Z")
_FILL_SUFFIX_RE = re.compile(r"-fill\Z")

_manifest_cache: dict[str, dict[str, object]] | None = None


def icon_render_mode(provider_id: str) -> Literal["fill", "stroke"]:
    """Internal freeform-engine lookup: ph fills, lu strokes."""
    try:
        return _SET_MODE[provider_id]
    except KeyError:
        raise InvalidArgumentError(f"unknown icon provider {provider_id!r}") from None


def _manifest() -> dict[str, dict[str, object]]:
    global _manifest_cache
    if _manifest_cache is None:
        with Path(_ICONS_DIR / "manifest.json").open(encoding="utf-8") as f:
            _manifest_cache = json.load(f)
    return _manifest_cache


def _provider_meta(provider_id: str) -> AssetProviderMeta:
    entry = _manifest()[_SET_DIRS[provider_id]]
    return AssetProviderMeta(
        provider_id=provider_id,
        license=str(entry["license"]),
        source_url=str(entry["source"]),
        source_commit=str(entry["commit"]),
        attribution=None,
        redistributable=True,
    )


def _canonical_name(provider_id: str, file_name: str) -> str:
    """Map a vendored filename to the canonical asset name.

    Phosphor fill files are `<name>-fill.svg`; the `-fill` weight suffix is not
    part of the public name (users write `ph:check`, not `ph:check-fill`).
    """
    name = file_name[: -len(".svg")]
    if provider_id == "ph" and _FILL_SUFFIX_RE.search(name):
        name = name[: -len("-fill")]
    return name


class IconProvider:
    """AssetProvider over one vendored icon set."""

    def __init__(self, provider_id: str) -> None:
        if provider_id not in _SET_DIRS:
            raise InvalidArgumentError(f"unknown icon provider {provider_id!r}")
        self.provider_id = provider_id
        self.kinds: frozenset[AssetKind] = frozenset({"icon"})
        self.provider_meta = _provider_meta(provider_id)

    def _dir(self) -> Path:
        return _ICONS_DIR / _SET_DIRS[self.provider_id]

    def _file_for(self, name: str) -> Path:
        if not _NAME_RE.match(name):
            raise InvalidArgumentError(f"invalid icon name {name!r}")
        if self.provider_id == "ph" and not _FILL_SUFFIX_RE.search(name):
            name = f"{name}-fill"
        return self._dir() / f"{name}.svg"

    def _read_svg(self, name: str) -> tuple[str, tuple[float, float, float, float]]:
        path = self._file_for(name)
        if not path.exists():
            raise InvalidArgumentError(f"icon {self.provider_id}:{name} not found")
        svg = path.read_text(encoding="utf-8")
        root = parse_svg(svg)
        vb = root.get("viewBox")
        if vb is None:
            raise InvalidArgumentError(f"icon {self.provider_id}:{name} has no viewBox")
        # SVG 规范允许逗号/空白任意混用作坐标分隔（viewBox="0,0 24,24"）
        parts = [p for p in re.split(r"[,\s]+", vb) if p]
        if len(parts) != 4:
            raise InvalidArgumentError(f"icon {self.provider_id}:{name} viewBox malformed: {vb!r}")
        try:
            x, y, w, h = (float(p) for p in parts)
            view_box = (x, y, w, h)
        except ValueError:
            raise InvalidArgumentError(
                f"icon {self.provider_id}:{name} viewBox malformed: {vb!r}"
            ) from None
        return svg, view_box

    # -- AssetProvider contract -------------------------------------------

    def search(
        self, query: str, *, kind: AssetKind | None = None, limit: int = 20
    ) -> list[AssetMeta]:
        if kind is not None and kind != "icon":
            return []
        metas: list[AssetMeta] = []
        needle = query.lower()
        for path in self._dir().glob("*.svg"):
            name = _canonical_name(self.provider_id, path.name)
            if needle and needle not in name.lower():
                continue
            ref = AssetRef(self.provider_id, "icon", name)
            metas.append(AssetMeta(ref=ref, title=name, tags=tuple(name.split("-"))))
        # sort by canonical name (not filename: the ph `-fill` weight suffix
        # would otherwise put `x-down` before `x`), then cap the global limit.
        metas.sort(key=lambda m: m.ref.name)
        return metas[:limit]

    def resolve(self, request: AssetRequest) -> ResolvedAsset:
        if request.params:
            raise InvalidArgumentError(f"icon provider {self.provider_id!r} does not accept params")
        ref = request.ref
        if ref.kind != "icon":
            raise InvalidArgumentError(
                f"icon provider {self.provider_id!r} does not support kind {ref.kind!r}"
            )
        svg, view_box = self._read_svg(ref.name)
        meta = AssetMeta(ref=ref, title=ref.name, tags=tuple(ref.name.split("-")))
        payload = SvgPayload(svg=svg, render_mode="freeform_svg", view_box=view_box)
        return ResolvedAsset(
            request=request, meta=meta, provider_meta=self.provider_meta, payload=payload
        )
