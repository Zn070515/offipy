"""offipy.assets.patterns.rings — deterministic concentric rings (A4).

``count`` full circles with evenly spaced radii that overscan the canvas; the
seed picks a center inside the safe central 30–70% region so the largest ring
stays within the accepted [-500, 1500] extent. ``thickness`` maps to the
stroke width.
"""

from __future__ import annotations

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open

_MAX_RADIUS = 790.0


def build(*, seed: int, background: str, count: int, thickness: float) -> str:
    stroke_width = 1.0 + thickness * 10.0
    rng = _Rng(seed)
    cx = rng.uniform(300.0, 700.0)
    cy = rng.uniform(300.0, 700.0)
    body: list[str] = []
    if background != "transparent":
        body.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    for i in range(count):
        radius = _MAX_RADIUS * (i + 1) / count
        body.append(
            f'<circle cx="{_fmt(cx)}" cy="{_fmt(cy)}" r="{_fmt(radius)}" '
            f'fill="none" stroke="{FG}" stroke-width="{_fmt(stroke_width)}"/>'
        )
    return svg_open() + "".join(body) + "</svg>"
