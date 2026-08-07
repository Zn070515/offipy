"""offipy.assets.primitives.process_arrow — editable step flow (A5).

Horizontal: N chevrons tiling the rect, each with editable step text. Vertical:
N rounded bands tiling the rect with a decorative down arrow between them. All
steps share the resolved ``fill`` (default transparent) with an accent outline
and ink text. Labels shrink to fit; below the 10pt readable floor the renderer
raises rather than producing unreadable output.
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE

from offipy.assets.primitives._common import (
    add_shape,
    fit_font_size_wrapped,
    require_rect,
    resolve_native_colors,
    set_shape_text,
    shape_elements,
)

_MIN_PT = 10.0
_PAD_FRACTION = 0.04
_TEXT_W_FRACTION = 0.72  # chevron text avoids the arrowhead
_ARROW_W_FRACTION = 0.08


def _step_font(step: str, w_px: float, h_px: float, start_pt: float) -> float:
    pad = max(_PAD_FRACTION * min(w_px, h_px), 2.0)
    return fit_font_size_wrapped(
        step, w_px - 2 * pad, h_px - 2 * pad, start_pt=start_pt, min_pt=_MIN_PT
    )


def _horizontal(slide, steps, colors, rect) -> list:
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    n = len(steps)
    cw = w / n
    shapes = []
    for i, step in enumerate(steps):
        chevron = add_shape(
            slide,
            MSO_SHAPE.CHEVRON,
            x + i * cw,
            y,
            cw,
            h,
            fill=colors["fill"],
            line=colors["accent"],
        )
        font_pt = _step_font(step, cw * _TEXT_W_FRACTION, h, h * 0.5)
        set_shape_text(chevron, step, font_size_pt=font_pt, color=colors["ink"])
        shapes.append(chevron)
    return shapes


def _vertical(slide, steps, colors, rect) -> list:
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    n = len(steps)
    hh = h / n
    arrow_w = w * _ARROW_W_FRACTION
    arrow_h = min(hh * 0.5, 14.0)
    shapes = []
    for i, step in enumerate(steps):
        band_y = y + i * hh
        band = add_shape(
            slide,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x,
            band_y,
            w,
            hh,
            fill=colors["fill"],
            line=colors["accent"],
        )
        font_pt = _step_font(step, w, hh, hh * 0.6)
        set_shape_text(band, step, font_size_pt=font_pt, color=colors["ink"])
        shapes.append(band)
        if i < n - 1:
            arrow = add_shape(
                slide,
                MSO_SHAPE.DOWN_ARROW,
                x + (w - arrow_w) / 2,
                band_y + hh - arrow_h / 2,
                arrow_w,
                arrow_h,
                fill=colors["accent"],
                line="transparent",
            )
            shapes.append(arrow)
    return shapes


def render(slide, params, context) -> tuple[object, ...]:
    """Draw the process arrow and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    steps = [s for s in params["steps"].split(",")]
    if params.get("direction", "horizontal") == "vertical":
        shapes = _vertical(slide, steps, colors, rect)
    else:
        shapes = _horizontal(slide, steps, colors, rect)
    return shape_elements(shapes)
