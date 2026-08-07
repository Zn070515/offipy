"""offipy.assets.patterns.dot_grid — deterministic dot grid (A4).

Dots on an axis-aligned grid; the seed offsets the grid origin within one
cell (no per-dot jitter, so the grid character holds). ``spacing`` maps to a
50..140px cell, ``radius`` to a dot radius capped at cell/3. A zero radius
intentionally produces an empty foreground (valid SVG, no dot elements).
"""

from __future__ import annotations

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open


def _cell(spacing: float) -> float:
    return 50.0 + (spacing - 0.5) * 60.0


def build(*, seed: int, background: str, spacing: float, radius: float) -> str:
    cell = _cell(spacing)
    rng = _Rng(seed)
    body: list[str] = []
    if background != "transparent":
        body.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    if radius > 0.0:
        dot_radius = min(radius * 40.0, cell / 3.0)
        off_x = rng.unit() * cell
        off_y = rng.unit() * cell
        x = off_x
        while x <= 1000.0:
            y = off_y
            while y <= 1000.0:
                body.append(
                    f'<circle cx="{_fmt(x)}" cy="{_fmt(y)}" r="{_fmt(dot_radius)}" fill="{FG}"/>'
                )
                y += cell
            x += cell
    return svg_open() + "".join(body) + "</svg>"
