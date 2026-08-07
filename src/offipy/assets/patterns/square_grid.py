"""offipy.assets.patterns.square_grid — deterministic square grid (A4).

Full-bleed vertical and horizontal lines on an axis-aligned grid; the seed
offsets the grid origin within one cell and geometry stays rectilinear.
``spacing`` maps to a 50..140px cell, ``thickness`` to a 1..9px stroke width.
Line positions are deduped on their formatted value so rounding can never
produce duplicate identical lines.
"""

from __future__ import annotations

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open


def _cell(spacing: float) -> float:
    return 50.0 + (spacing - 0.5) * 60.0


def _line_positions(offset: float, cell: float) -> list[str]:
    seen: set[str] = set()
    positions: list[str] = []
    pos = offset
    while pos <= 1000.0:
        text = _fmt(pos)
        if text not in seen:
            seen.add(text)
            positions.append(text)
        pos += cell
    return positions


def build(*, seed: int, background: str, spacing: float, thickness: float) -> str:
    cell = _cell(spacing)
    line_width = 1.0 + thickness * 8.0
    rng = _Rng(seed)
    off_x = rng.unit() * cell
    off_y = rng.unit() * cell
    body: list[str] = []
    if background != "transparent":
        body.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    for x in _line_positions(off_x, cell):
        body.append(
            f'<line x1="{x}" y1="0" x2="{x}" y2="1000" stroke="{FG}" '
            f'stroke-width="{_fmt(line_width)}"/>'
        )
    for y in _line_positions(off_y, cell):
        body.append(
            f'<line x1="0" y1="{y}" x2="1000" y2="{y}" stroke="{FG}" '
            f'stroke-width="{_fmt(line_width)}"/>'
        )
    return svg_open() + "".join(body) + "</svg>"
