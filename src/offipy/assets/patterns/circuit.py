"""offipy.assets.patterns.circuit — deterministic orthogonal circuit board (A4).

``nodes`` distinct cells are drawn from a coarse 10x10 logical grid via a
partial Fisher-Yates shuffle and sorted row-major so index-adjacent nodes sit
close together. Nodes are joined with Manhattan L-routes (horizontal/vertical
segments only); ``density`` scales the number of skip-ahead extra edges up to a
cap of ``2 * nodes``. A small per-node jitter keeps the grid from reading as a
perfect lattice while positions stay deterministic.
"""

from __future__ import annotations

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open

_GRID_W = 10
_GRID_H = 10
_CELL = 100.0
_MARGIN = 50.0
_JITTER = 15.0
_MAX_EDGE_MULT = 2
_NODE_R = 5.0
_LINE_W = 2.0


def _route(x1: float, y1: float, x2: float, y2: float) -> str:
    if abs(x2 - x1) >= abs(y2 - y1):
        return f"M{_fmt(x1)} {_fmt(y1)} H{_fmt(x2)} V{_fmt(y2)}"
    return f"M{_fmt(x1)} {_fmt(y1)} V{_fmt(y2)} H{_fmt(x2)}"


def build(*, seed: int, background: str, nodes: int, density: float) -> str:
    """Return an SVG template for the circuit pattern."""
    total = _GRID_W * _GRID_H
    rng = _Rng(seed)
    cells = list(range(total))
    for i in range(nodes):
        j = i + rng.choice_index(total - i)
        cells[i], cells[j] = cells[j], cells[i]
    selected = sorted(cells[:nodes], key=lambda idx: (idx // _GRID_W, idx % _GRID_W))

    pts: list[tuple[float, float]] = []
    for idx in selected:
        col = idx % _GRID_W
        row = idx // _GRID_W
        cx = _MARGIN + col * _CELL + rng.uniform(-_JITTER, _JITTER)
        cy = _MARGIN + row * _CELL + rng.uniform(-_JITTER, _JITTER)
        pts.append((cx, cy))

    edges: list[tuple[int, int]] = [(i, i + 1) for i in range(nodes - 1)]
    extra = round(density * 2)
    max_edges = _MAX_EDGE_MULT * nodes
    for i in range(nodes - 1):
        for step in range(2, 2 + extra):
            j = i + step
            if j >= nodes or len(edges) >= max_edges:
                break
            edges.append((i, j))

    body: list[str] = []
    if background != "transparent":
        body.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    for i, j in edges:
        x1, y1 = pts[i]
        x2, y2 = pts[j]
        body.append(
            f'<path d="{_route(x1, y1, x2, y2)}" fill="none" stroke="{FG}" '
            f'stroke-width="{_fmt(_LINE_W)}"/>'
        )
    for cx, cy in pts:
        body.append(f'<circle cx="{_fmt(cx)}" cy="{_fmt(cy)}" r="{_fmt(_NODE_R)}" fill="{FG}"/>')
    return svg_open() + "".join(body) + "</svg>"
