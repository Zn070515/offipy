"""offipy.assets.providers.primitives — native presentation primitives provider (A5).

Owns the parameter schema and strict validation for the eight editable
PowerPoint-native primitives. ``resolve`` validates the request and wraps the
canonical validated string params in a ``NativeShapePayload``; it never touches
PowerPoint, theme, or the render rect (that is the renderer's job).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from offipy.assets.color import validate_color_value
from offipy.assets.model import (
    AssetKind,
    AssetMeta,
    AssetProviderMeta,
    AssetRef,
    AssetRequest,
    NativeShapePayload,
    ResolvedAsset,
)
from offipy.exceptions import InvalidArgumentError

_PROVIDER_ID = "primitives"
_KIND: AssetKind = "primitive"
_SOURCE_URL = "https://github.com/Zn070515/offipy"

_PRIMITIVE_ORDER = (
    "quote-mark",
    "section-number",
    "label-pill",
    "metric-badge",
    "timeline-node",
    "process-arrow",
    "device-frame",
    "browser-mockup",
)

_TITLES: dict[str, str] = {
    "quote-mark": "Quote Mark",
    "section-number": "Section Number",
    "label-pill": "Label Pill",
    "metric-badge": "Metric Badge",
    "timeline-node": "Timeline Node",
    "process-arrow": "Process Arrow",
    "device-frame": "Device Frame",
    "browser-mockup": "Browser Mockup",
}

_TAGS: dict[str, tuple[str, ...]] = {
    "quote-mark": ("quote", "testimonial", "text"),
    "section-number": ("number", "section", "label"),
    "label-pill": ("label", "pill", "tag", "chip"),
    "metric-badge": ("metric", "badge", "kpi", "delta"),
    "timeline-node": ("timeline", "node", "phase", "milestone"),
    "process-arrow": ("process", "arrow", "steps", "flow"),
    "device-frame": ("device", "phone", "tablet", "desktop", "frame"),
    "browser-mockup": ("browser", "mockup", "window", "url"),
}

# Common `fill` default per primitive; `accent` always defaults to the theme
# accent token. Freeze label-pill as accent-fill + bg-contrast text.
_FILL_DEFAULT: dict[str, str] = {
    "quote-mark": "transparent",
    "section-number": "transparent",
    "label-pill": "accent",
    "metric-badge": "surface",
    "timeline-node": "transparent",
    "process-arrow": "transparent",
    "device-frame": "surface",
    "browser-mockup": "surface",
}

_FORBIDDEN_PARAMS = frozenset({"screenshot", "src", "image", "image-src"})

_INT_RE = re.compile(r"-?[0-9]+\Z")


@dataclass(frozen=True)
class _TextSpec:
    required: bool
    max_len: int


@dataclass(frozen=True)
class _IntSpec:
    low: int
    high: int


@dataclass(frozen=True)
class _EnumSpec:
    choices: tuple[str, ...]
    default: str | None


@dataclass(frozen=True)
class _ListSpec:
    min_items: int
    max_items: int
    item_max_len: int


_ParamSpec = _TextSpec | _IntSpec | _EnumSpec | _ListSpec

# Canonical keys use hyphens. Common `accent` / `fill` are allowed for every
# primitive and handled by _coerce_params, not listed per-primitive.
_PRIMITIVE_SPECS: dict[str, dict[str, _ParamSpec]] = {
    "quote-mark": {"text": _TextSpec(required=True, max_len=240)},
    "section-number": {
        "number": _IntSpec(low=0, high=9999),
        "label": _TextSpec(required=False, max_len=120),
    },
    "label-pill": {"text": _TextSpec(required=True, max_len=120)},
    "metric-badge": {
        "value": _TextSpec(required=True, max_len=80),
        "label": _TextSpec(required=False, max_len=120),
        "delta": _TextSpec(required=False, max_len=40),
    },
    "timeline-node": {
        "label": _TextSpec(required=False, max_len=120),
        "phase": _EnumSpec(choices=("past", "current", "future"), default="current"),
    },
    "process-arrow": {
        "steps": _ListSpec(min_items=2, max_items=8, item_max_len=80),
        "direction": _EnumSpec(choices=("horizontal", "vertical"), default="horizontal"),
    },
    "device-frame": {
        "device": _EnumSpec(choices=("phone", "tablet", "desktop"), default=None),
    },
    "browser-mockup": {
        "title": _TextSpec(required=False, max_len=120),
        "url": _TextSpec(required=False, max_len=240),
    },
}


def _coerce_params(primitive: str, params: tuple[tuple[str, str], ...]) -> dict[str, str]:
    """Validate URI params into a canonical string param set with defaults.

    Returns a fully-resolved dict (common accent/fill + every defaulted spec
    param), so the payload the renderer receives is self-contained. Optional
    text params are omitted when absent.
    """
    spec = _PRIMITIVE_SPECS.get(primitive)
    if spec is None:
        raise InvalidArgumentError(
            f"unknown native primitive {primitive!r}; expected one of {list(_PRIMITIVE_SPECS)}"
        )
    allowed = set(spec) | {"accent", "fill"}
    given: dict[str, str] = {}
    for key, value in params:
        if key in _FORBIDDEN_PARAMS:
            raise InvalidArgumentError(f"param {key!r} is not supported in v0.14")
        if key not in allowed:
            raise InvalidArgumentError(
                f"unknown param {key!r} for primitive {primitive!r}; allowed: {sorted(allowed)}"
            )
        if key in given:
            raise InvalidArgumentError(f"duplicate param {key!r}")
        given[key] = value

    out: dict[str, str] = {
        "accent": validate_color_value(given.get("accent", "accent")),
        "fill": validate_color_value(given.get("fill", _FILL_DEFAULT[primitive])),
    }
    for key, item in spec.items():
        raw = given.get(key)
        if isinstance(item, _TextSpec):
            if raw is None:
                if item.required:
                    raise InvalidArgumentError(f"missing required param {key!r}")
                continue
            text = raw.strip()
            if item.required and not text:
                raise InvalidArgumentError(f"param {key!r} must be non-empty")
            if not text:
                continue  # optional empty → omit
            if len(text) > item.max_len:
                raise InvalidArgumentError(
                    f"param {key!r} exceeds max length {item.max_len}, got {len(text)}"
                )
            out[key] = text
        elif isinstance(item, _IntSpec):
            if raw is None:
                raise InvalidArgumentError(f"missing required param {key!r}")
            value = raw.strip()
            if not _INT_RE.match(value):
                raise InvalidArgumentError(f"param {key!r} must be an integer, got {raw!r}")
            parsed = int(value)
            if not item.low <= parsed <= item.high:
                raise InvalidArgumentError(
                    f"param {key!r} must be in [{item.low}, {item.high}], got {parsed}"
                )
            out[key] = str(parsed)
        elif isinstance(item, _EnumSpec):
            if raw is None:
                if item.default is None:
                    raise InvalidArgumentError(f"missing required param {key!r}")
                out[key] = item.default
                continue
            value = raw.strip()
            if value not in item.choices:
                raise InvalidArgumentError(
                    f"param {key!r} must be one of {list(item.choices)}, got {raw!r}"
                )
            out[key] = value
        elif isinstance(item, _ListSpec):
            if raw is None:
                raise InvalidArgumentError(f"missing required param {key!r}")
            items = [part.strip() for part in raw.split(",")]
            if len(items) < item.min_items or len(items) > item.max_items:
                raise InvalidArgumentError(
                    f"param {key!r} needs {item.min_items}..{item.max_items} items, "
                    f"got {len(items)}"
                )
            for part in items:
                if not part:
                    raise InvalidArgumentError(f"param {key!r} contains an empty item")
                if len(part) > item.item_max_len:
                    raise InvalidArgumentError(
                        f"param {key!r} item exceeds max length {item.item_max_len}"
                    )
            out[key] = ",".join(items)
    return out


class PrimitivesProvider:
    """AssetProvider over the eight editable native presentation primitives."""

    def __init__(self) -> None:
        self.provider_id = _PROVIDER_ID
        self.kinds: frozenset[AssetKind] = frozenset({"primitive"})
        self.provider_meta = AssetProviderMeta(
            provider_id=_PROVIDER_ID,
            license="MIT",
            source_url=_SOURCE_URL,
            source_commit=None,
            attribution=None,
            redistributable=True,
            first_party=True,
        )

    def search(
        self, query: str, *, kind: AssetKind | None = None, limit: int = 20
    ) -> list[AssetMeta]:
        if kind is not None and kind != "primitive":
            return []
        needle = query.lower()
        metas: list[AssetMeta] = []
        for name in _PRIMITIVE_ORDER:
            tags = _TAGS[name]
            haystack = " ".join((name, _TITLES[name], *tags)).lower()
            if needle and needle not in haystack:
                continue
            ref = AssetRef(_PROVIDER_ID, "primitive", name)
            metas.append(AssetMeta(ref=ref, title=_TITLES[name], tags=tags))
        return metas[:limit]

    def resolve(self, request: AssetRequest) -> ResolvedAsset:
        ref = request.ref
        if ref.kind != "primitive":
            raise InvalidArgumentError(f"primitives provider does not support kind {ref.kind!r}")
        params = _coerce_params(ref.name, request.params)
        payload = NativeShapePayload(primitive=ref.name, params=tuple(params.items()))
        meta = AssetMeta(ref=ref, title=_TITLES[ref.name], tags=_TAGS[ref.name])
        return ResolvedAsset(
            request=request, meta=meta, provider_meta=self.provider_meta, payload=payload
        )
