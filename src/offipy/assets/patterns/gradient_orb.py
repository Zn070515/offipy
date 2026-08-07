"""offipy.assets.patterns.gradient_orb — deterministic soft radial glows (A4).

Each orb is a ``<circle>`` filled with its own ``<radialGradient>`` that fades
the foreground color to opacity 0. ``blur`` spreads the gradient stops: at 0 the
core stays opaque to ~85% of the radius (hard falloff), at 1 it starts fading at
~25% (wide soft halo). Only gradient + opacity features covered by the A1 probe
are used; no SVG ``filter``/``feGaussianBlur``.
"""

from __future__ import annotations

from offipy.assets.patterns._common import BG, FG, _fmt, _Rng, svg_open

_CENTER_LOW = 200.0
_CENTER_HIGH = 800.0
_RADIUS_LOW = 100.0
_RADIUS_HIGH = 220.0


def _falloff(blur: float) -> float:
    return 0.85 - blur * 0.6


def build(*, seed: int, background: str, orb_count: int, blur: float) -> str:
    """Return an SVG template for the gradient-orb pattern."""
    falloff = _falloff(blur)
    rng = _Rng(seed)
    orbs: list[tuple[float, float, float]] = []
    for _ in range(orb_count):
        cx = rng.uniform(_CENTER_LOW, _CENTER_HIGH)
        cy = rng.uniform(_CENTER_LOW, _CENTER_HIGH)
        radius = rng.uniform(_RADIUS_LOW, _RADIUS_HIGH)
        orbs.append((cx, cy, radius))

    body: list[str] = []
    if background != "transparent":
        body.append(f'<rect width="1000" height="1000" fill="{BG}"/>')
    defs = ["<defs>"]
    for i in range(orb_count):
        defs.append(
            f'<radialGradient id="orb-{i}" cx="50%" cy="50%" r="50%">'
            f'<stop offset="0%" stop-color="{FG}" stop-opacity="1"/>'
            f'<stop offset="{_fmt(falloff * 100.0)}%" stop-color="{FG}" stop-opacity="1"/>'
            f'<stop offset="100%" stop-color="{FG}" stop-opacity="0"/>'
            f"</radialGradient>"
        )
    defs.append("</defs>")
    body.append("".join(defs))
    for i, (cx, cy, radius) in enumerate(orbs):
        body.append(
            f'<circle cx="{_fmt(cx)}" cy="{_fmt(cy)}" r="{_fmt(radius)}" fill="url(#orb-{i})"/>'
        )
    return svg_open() + "".join(body) + "</svg>"
