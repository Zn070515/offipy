"""offipy.assets.patterns._common — deterministic procedural SVG foundation (A4).

Shared primitives for the eight procedural patterns: a fixed 32-bit PRNG
(SplitMix32) whose arithmetic is masked to 32 bits so output is byte-identical
across Python versions, a stable float formatter, deterministic SVG root
serialization (viewBox normalized to 1000×1000), and the theme color
sentinels that flow through ``SvgTemplatePayload.color_slots``.
"""

from __future__ import annotations

import math

from offipy.exceptions import InvalidArgumentError

# Theme color sentinels. These appear literally in a pattern template and are
# swapped for resolved colors by materialize_svg_template via color_slots.
FG = "__OFFIPY_ASSET_FG__"
BG = "__OFFIPY_ASSET_BG__"

_SVG_NS = "http://www.w3.org/2000/svg"
_VIEW_BOX = "0 0 1000 1000"

_MASK32 = 0xFFFFFFFF
_SPLITMIX_GAMMA = 0x9E3779B9  # golden ratio; splitmix step constant
_SPLITMIX_M1 = 0x85EBCA6B
_SPLITMIX_M2 = 0xC2B2AE35


class _Rng:
    """SplitMix32 deterministic PRNG, output masked to 32 bits.

    Exists so pattern output is byte-identical across Python 3.10–3.13 and
    does not depend on the Python ``random`` implementation. The state is a
    single unsigned 32-bit counter advanced by the golden ratio; each output
    applies the fmix32 finalizer.
    """

    __slots__ = ("_state",)

    def __init__(self, seed: int) -> None:
        self._state = seed & _MASK32

    def u32(self) -> int:
        self._state = (self._state + _SPLITMIX_GAMMA) & _MASK32
        z = self._state
        z = ((z ^ (z >> 16)) * _SPLITMIX_M1) & _MASK32
        z = ((z ^ (z >> 13)) * _SPLITMIX_M2) & _MASK32
        return (z ^ (z >> 16)) & _MASK32

    def unit(self) -> float:
        """Uniform float in [0, 1)."""
        return self.u32() / 0x100000000

    def uniform(self, a: float, b: float) -> float:
        """Uniform float in [a, b)."""
        return a + (b - a) * self.unit()

    def choice_index(self, n: int) -> int:
        """Uniform integer in [0, n)."""
        if n <= 0:
            raise InvalidArgumentError(f"choice_index expects positive n, got {n}")
        return self.u32() % n


def _fmt(value: float, *, digits: int = 3) -> str:
    """Format a float for SVG output deterministically.

    Round to at most ``digits`` decimals (default 3), normalize ``-0`` to
    ``0``, trim trailing zeros and the decimal point, and reject non-finite
    input so a bad coordinate can never slip a NaN/Inf into the XML.
    """
    if not math.isfinite(value):
        raise InvalidArgumentError(f"cannot format non-finite number: {value!r}")
    rounded = round(value, digits)
    if rounded == 0:
        rounded = 0.0
    text = f"{rounded:.{digits}f}"
    return text.rstrip("0").rstrip(".")


def svg_open(extra_attrs: tuple[tuple[str, str], ...] = ()) -> str:
    """Deterministic SVG root opening tag with a normalized 1000×1000 viewBox.

    Attribute order is fixed so template bytes are stable. Patterns append
    their body and ``</svg>`` themselves.
    """
    attrs = [f'xmlns="{_SVG_NS}"', f'viewBox="{_VIEW_BOX}"']
    attrs.extend(f'{key}="{value}"' for key, value in extra_attrs)
    return "<svg " + " ".join(attrs) + ">"
