"""offipy.assets.patterns.topography — deterministic contour-like field (A4).

``lines`` wavy polylines run from left edge to right edge, each baseline evenly
spaced down the canvas. Vertical displacement is a sum of three sinusoidal
components whose amplitudes/frequencies/phases come from the seed-driven PRNG;
``density`` scales sample density, amplitude, frequency, and stroke width so the
texture gets richer but stays bounded. Paths are smoothed with uniform
Catmull-Rom, the same recipe as the wave pattern.
"""

from __future__ import annotations

import math

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open

_COMPONENTS = 3
_SAMPLES_LOW = 40
_SAMPLES_HIGH = 80
_STEP = 1000.0


def _line_path(
    baseline: float, samples: int, amp_scale: float, freq_scale: float, rng: _Rng
) -> str:
    comps: list[tuple[float, float, float]] = []
    for _ in range(_COMPONENTS):
        amp = rng.uniform(8.0, 18.0) * amp_scale
        freq = rng.uniform(1.0, 3.0) * freq_scale
        phase = rng.unit() * 2.0 * math.pi
        comps.append((amp, freq, phase))

    def y_at(x: float) -> float:
        return baseline + sum(
            amp * math.sin(2.0 * math.pi * freq * x / _STEP + phase) for amp, freq, phase in comps
        )

    step = _STEP / samples
    xs = [-step] + [i * step for i in range(samples + 1)] + [_STEP + step]
    ys = [y_at(x) for x in xs]
    parts = [f"M{_fmt(0.0)} {_fmt(ys[1])}"]
    for i in range(samples):
        x0, y0 = xs[i + 1], ys[i + 1]
        x1, y1 = xs[i + 2], ys[i + 2]
        xp, yp = xs[i], ys[i]
        x2, y2 = xs[i + 3], ys[i + 3]
        c1 = f"{_fmt(x0 + (x1 - xp) / 6.0)} {_fmt(y0 + (y1 - yp) / 6.0)}"
        c2 = f"{_fmt(x1 - (x2 - x0) / 6.0)} {_fmt(y1 - (y2 - y0) / 6.0)}"
        parts.append(f"C{c1} {c2} {_fmt(x1)} {_fmt(y1)}")
    return " ".join(parts)


def build(*, seed: int, background: str, density: float, lines: int) -> str:
    """Return an SVG template for the topography pattern."""
    samples = _SAMPLES_LOW + round(density * (_SAMPLES_HIGH - _SAMPLES_LOW))
    amp_scale = 0.5 + density * 1.0
    freq_scale = 1.0 + density * 0.8
    stroke_width = 1.0 + density * 2.0
    rng = _Rng(seed)
    parts: list[str] = []
    if background != "transparent":
        parts.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    for i in range(lines):
        baseline = _STEP * (i + 1) / (lines + 1)
        parts.append(
            f'<path d="{_line_path(baseline, samples, amp_scale, freq_scale, rng)}" '
            f'fill="none" stroke="{FG}" stroke-width="{_fmt(stroke_width)}"/>'
        )
    return svg_open() + "".join(parts) + "</svg>"
