"""offipy.assets.patterns.wave — deterministic horizontal wave lines (A4).

Layered-sinusoid waves sampled on a fixed grid and turned into smooth cubic
Bézier paths via uniform Catmull-Rom. Line count derives from ``density``,
stroke width from ``thickness``; the seed drives per-line phase/amplitude
perturbation within bounded ranges so coordinates stay inside the modest
overscan allowed by the acceptance tests.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open

_SEGMENTS = 16
_MARGIN = 80.0


@dataclass(frozen=True)
class _Wave:
    base: float
    amp1: float
    wl1: float
    ph1: float
    amp2: float
    wl2: float
    ph2: float

    def y(self, x: float) -> float:
        return (
            self.base
            + self.amp1 * math.sin(2.0 * math.pi * x / self.wl1 + self.ph1)
            + self.amp2 * math.sin(2.0 * math.pi * x / self.wl2 + self.ph2)
        )


def _wave_path(wave: _Wave) -> str:
    """One wave as a single cubic-Bézier path across the 1000px canvas."""
    step = 1000.0 / _SEGMENTS
    xs = [-step] + [i * step for i in range(_SEGMENTS + 1)] + [1000.0 + step]
    ys = [wave.y(x) for x in xs]
    parts = [f"M{_fmt(0.0)} {_fmt(ys[1])}"]
    for i in range(_SEGMENTS):
        x0, y0 = xs[i + 1], ys[i + 1]
        x1, y1 = xs[i + 2], ys[i + 2]
        xp, yp = xs[i], ys[i]
        x2, y2 = xs[i + 3], ys[i + 3]
        c1 = f"{_fmt(x0 + (x1 - xp) / 6.0)} {_fmt(y0 + (y1 - yp) / 6.0)}"
        c2 = f"{_fmt(x1 - (x2 - x0) / 6.0)} {_fmt(y1 - (y2 - y0) / 6.0)}"
        parts.append(f"C{c1} {c2} {_fmt(x1)} {_fmt(y1)}")
    return " ".join(parts)


def build(*, seed: int, background: str, density: float, thickness: float) -> str:
    """Return an SVG template for the wave pattern."""
    line_count = 2 + round(density * 8)
    stroke_width = 1.5 + thickness * 10
    rng = _Rng(seed)
    parts: list[str] = []
    if background != "transparent":
        parts.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    for i in range(line_count):
        base = _MARGIN + i * ((1000.0 - 2 * _MARGIN) / (line_count - 1))
        amp1 = rng.uniform(15.0, 35.0)
        wl1 = rng.uniform(200.0, 400.0)
        ph1 = rng.unit() * 2.0 * math.pi
        amp2 = amp1 * rng.uniform(0.0, 0.35)
        wl2 = rng.uniform(100.0, 250.0)
        ph2 = rng.unit() * 2.0 * math.pi
        wave = _Wave(base, amp1, wl1, ph1, amp2, wl2, ph2)
        parts.append(
            f'<path d="{_wave_path(wave)}" fill="none" stroke="{FG}" '
            f'stroke-width="{_fmt(stroke_width)}"/>'
        )
    return svg_open() + "".join(parts) + "</svg>"
