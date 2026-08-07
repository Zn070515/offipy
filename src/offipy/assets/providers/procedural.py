"""offipy.assets.providers.procedural — deterministic pattern provider (A4).

Owns the parameter schema and strict validation for the eight procedural
patterns. Generation lives in ``offipy.assets.patterns.*`` and is wired into
``resolve`` in A4 Task 10; until then resolve validates and raises, so the
provider can be developed and tested schema-first.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass

from offipy.assets.color import validate_color_value
from offipy.assets.model import (
    AssetKind,
    AssetMeta,
    AssetProviderMeta,
    AssetRef,
    AssetRequest,
    ResolvedAsset,
)
from offipy.exceptions import InvalidArgumentError

_PROVIDER_ID = "procedural"
_KIND: AssetKind = "pattern"
_SOURCE_URL = "https://github.com/Zn070515/offipy"

_PATTERN_ORDER = (
    "wave",
    "blob",
    "dot-grid",
    "square-grid",
    "rings",
    "topography",
    "circuit",
    "gradient-orb",
)

_PATTERN_TITLES: dict[str, str] = {
    "wave": "Wave Lines",
    "blob": "Organic Blob",
    "dot-grid": "Dot Grid",
    "square-grid": "Square Grid",
    "rings": "Concentric Rings",
    "topography": "Topography Contours",
    "circuit": "Circuit Network",
    "gradient-orb": "Gradient Orbs",
}

_PATTERN_TAGS: dict[str, tuple[str, ...]] = {
    "wave": ("background", "line", "wave"),
    "blob": ("background", "organic", "shape"),
    "dot-grid": ("background", "grid", "dot"),
    "square-grid": ("background", "grid", "line"),
    "rings": ("background", "concentric", "ring"),
    "topography": ("background", "line", "contour"),
    "circuit": ("background", "tech", "circuit"),
    "gradient-orb": ("background", "orb", "glow"),
}

# Frozen shared params (rev1.2 §3.2): seed / foreground / background.
_DEFAULT_FOREGROUND = "accent"
_DEFAULT_BACKGROUND = "transparent"
_SEED_LOW = -(1 << 31)
_SEED_HIGH = (1 << 31) - 1


@dataclass(frozen=True)
class _IntParam:
    low: int
    high: int
    default: int


@dataclass(frozen=True)
class _FloatParam:
    low: float
    high: float
    default: float


# Canonical keys use hyphens (orb-count, not orb_count).
_PATTERN_PARAMS: dict[str, dict[str, _IntParam | _FloatParam]] = {
    "wave": {
        "density": _FloatParam(0.0, 1.0, 0.5),
        "thickness": _FloatParam(0.0, 1.0, 0.5),
    },
    "blob": {"complexity": _IntParam(2, 8, 5)},
    "dot-grid": {
        "spacing": _FloatParam(0.5, 2.0, 1.0),
        "radius": _FloatParam(0.0, 0.5, 0.15),
    },
    "square-grid": {
        "spacing": _FloatParam(0.5, 2.0, 1.0),
        "thickness": _FloatParam(0.0, 1.0, 0.5),
    },
    "rings": {
        "count": _IntParam(1, 12, 5),
        "thickness": _FloatParam(0.0, 1.0, 0.5),
    },
    "topography": {
        "density": _FloatParam(0.0, 1.0, 0.5),
        "lines": _IntParam(3, 24, 10),
    },
    "circuit": {
        "nodes": _IntParam(4, 60, 20),
        "density": _FloatParam(0.0, 1.0, 0.5),
    },
    "gradient-orb": {
        "orb-count": _IntParam(1, 6, 3),
        "blur": _FloatParam(0.0, 1.0, 0.5),
    },
}

_INT_RE = re.compile(r"-?[0-9]+\Z")
_FLOAT_RE = re.compile(r"-?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)\Z")


def _parse_int(value: str, name: str, low: int, high: int) -> int:
    if not _INT_RE.match(value):
        raise InvalidArgumentError(f"param {name!r} must be an integer, got {value!r}")
    parsed = int(value)
    if not low <= parsed <= high:
        raise InvalidArgumentError(f"param {name!r} must be in [{low}, {high}], got {parsed}")
    return parsed


def _parse_float(value: str, name: str, low: float, high: float) -> float:
    if not _FLOAT_RE.match(value):
        raise InvalidArgumentError(f"param {name!r} must be a finite decimal, got {value!r}")
    parsed = float(value)
    if not math.isfinite(parsed):
        raise InvalidArgumentError(f"param {name!r} must be a finite decimal, got {value!r}")
    if not low <= parsed <= high:
        raise InvalidArgumentError(f"param {name!r} must be in [{low}, {high}], got {parsed}")
    return parsed


def _coerce_params(pattern: str, params: tuple[tuple[str, str], ...]) -> dict[str, object]:
    """Validate URI params into a typed canonical param set with defaults."""
    spec = _PATTERN_PARAMS.get(pattern)
    if spec is None:
        raise InvalidArgumentError(
            f"unknown procedural pattern {pattern!r}; expected one of {list(_PATTERN_PARAMS)}"
        )
    allowed = set(spec) | {"seed", "foreground", "background"}
    given: dict[str, str] = {}
    for key, value in params:
        if key not in allowed:
            raise InvalidArgumentError(
                f"unknown param {key!r} for pattern {pattern!r}; allowed: {sorted(allowed)}"
            )
        if key in given:
            raise InvalidArgumentError(f"duplicate param {key!r}")
        given[key] = value
    typed: dict[str, object] = {
        "seed": _parse_int(given.get("seed", "0"), "seed", _SEED_LOW, _SEED_HIGH),
    }
    for key, param_spec in spec.items():
        raw = given.get(key, str(param_spec.default))
        if isinstance(param_spec, _IntParam):
            typed[key] = _parse_int(raw, key, param_spec.low, param_spec.high)
        else:
            typed[key] = _parse_float(raw, key, param_spec.low, param_spec.high)
    typed["foreground"] = validate_color_value(given.get("foreground", _DEFAULT_FOREGROUND))
    typed["background"] = validate_color_value(given.get("background", _DEFAULT_BACKGROUND))
    return typed


class ProceduralProvider:
    """AssetProvider over the eight deterministic procedural patterns."""

    def __init__(self) -> None:
        self.provider_id = _PROVIDER_ID
        self.kinds: frozenset[AssetKind] = frozenset({"pattern"})
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
        if kind is not None and kind != "pattern":
            return []
        needle = query.lower()
        metas: list[AssetMeta] = []
        for name in _PATTERN_ORDER:
            tags = _PATTERN_TAGS[name]
            haystack = " ".join((name, _PATTERN_TITLES[name], *tags)).lower()
            if needle and needle not in haystack:
                continue
            ref = AssetRef(_PROVIDER_ID, "pattern", name)
            metas.append(AssetMeta(ref=ref, title=_PATTERN_TITLES[name], tags=tags))
        return metas[:limit]

    def resolve(self, request: AssetRequest) -> ResolvedAsset:
        ref = request.ref
        if ref.kind != "pattern":
            raise InvalidArgumentError(f"procedural provider does not support kind {ref.kind!r}")
        _coerce_params(ref.name, request.params)
        # Generation + SvgTemplatePayload construction lands in A4 Task 10.
        raise NotImplementedError(
            f"procedural pattern {ref.name!r} generation is wired in A4 Task 10"
        )
