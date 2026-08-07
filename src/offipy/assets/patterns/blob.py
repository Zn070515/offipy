"""offipy.assets.patterns.blob — deterministic organic blob (A4).

One closed, smooth cubic path centered on the canvas. The ``complexity``
parameter maps to the number of radial anchors: complexity 2 maps to the
8-anchor minimum (fewer anchors produce an unrecognizable pinch), and
complexity 8 maps to 16 anchors. The seed jitters each anchor's angle (kept
within half the inter-anchor gap so anchors never swap) and radial distance;
uniform Catmull-Rom turns the anchor ring into a closed, non-exploding path.
"""

from __future__ import annotations

import math

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open

_CENTER = 500.0
_BASE_RADIUS = 280.0
_MIN_ANCHORS = 8


def _blob_path(rng: _Rng, anchors: int) -> str:
    step = 2.0 * math.pi / anchors
    points: list[tuple[float, float]] = []
    for i in range(anchors):
        theta = i * step + rng.uniform(-step / 2.0, step / 2.0)
        radius = _BASE_RADIUS * rng.uniform(0.7, 1.3)
        points.append((_CENTER + radius * math.cos(theta), _CENTER + radius * math.sin(theta)))
    parts = [f"M{_fmt(points[0][0])} {_fmt(points[0][1])}"]
    for i in range(anchors):
        p0 = points[i]
        p1 = points[(i + 1) % anchors]
        pm = points[(i - 1) % anchors]
        p2 = points[(i + 2) % anchors]
        c1 = (p0[0] + (p1[0] - pm[0]) / 6.0, p0[1] + (p1[1] - pm[1]) / 6.0)
        c2 = (p1[0] - (p2[0] - p0[0]) / 6.0, p1[1] - (p2[1] - p0[1]) / 6.0)
        parts.append(
            f"C{_fmt(c1[0])} {_fmt(c1[1])} {_fmt(c2[0])} {_fmt(c2[1])} {_fmt(p1[0])} {_fmt(p1[1])}"
        )
    return " ".join(parts)


def build(*, seed: int, background: str, complexity: int) -> str:
    """Return an SVG template for the blob pattern."""
    anchors = max(_MIN_ANCHORS, complexity * 2)
    rng = _Rng(seed)
    body: list[str] = []
    if background != "transparent":
        body.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    body.append(f'<path d="{_blob_path(rng, anchors)}" fill="{FG}"/>')
    return svg_open() + "".join(body) + "</svg>"
