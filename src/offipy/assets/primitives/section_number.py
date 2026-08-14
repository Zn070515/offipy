"""offipy.assets.primitives.section_number — editable section number (A5).

Composition: a short accent bar on the left, the canonical decimal number as a
bold editable textbox, and an optional wrapping label below it. The number is
a single line; the label wraps inside the rect.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.shapes import MSO_SHAPE

from offipy.assets.primitives._common import (
    add_shape,
    add_textbox,
    fit_font_size,
    fit_font_size_wrapped,
    require_rect,
    resolve_native_colors,
    shape_elements,
)

if TYPE_CHECKING:
    from collections.abc import Mapping

    from offipy.assets.model import AssetRenderContext

_BAR_W_FRACTION = 0.05
_PAD_FRACTION = 0.06
_NUM_H_FRACTION = 0.55
_MIN_PT = 8.0
_LABEL_START_PT = 18.0


def render(
    slide: Any,
    params: Mapping[str, str],
    context: AssetRenderContext,
) -> tuple[object, ...]:
    """Draw the section number and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    pad = max(_PAD_FRACTION * min(w, h), 4.0)

    shapes = []
    bar_w = w * _BAR_W_FRACTION
    bar = add_shape(
        slide,
        MSO_SHAPE.RECTANGLE,
        x + pad,
        y + pad,
        bar_w,
        h - 2 * pad,
        fill=colors["accent"],
        line="transparent",
    )
    shapes.append(bar)

    body_x = x + pad + bar_w + pad
    body_w = w - 2 * pad - bar_w - pad
    num_h = h * _NUM_H_FRACTION
    num_pt = fit_font_size(params["number"], body_w, num_h, start_pt=num_h * 0.8, min_pt=_MIN_PT)
    number = add_textbox(
        slide,
        body_x,
        y + pad,
        body_w,
        num_h,
        text=params["number"],
        font_size_pt=num_pt,
        color=colors["ink"],
        bold=True,
    )
    shapes.append(number)

    label = params.get("label")
    if label:
        label_y = y + pad + num_h
        label_h = h - 2 * pad - num_h
        label_pt = fit_font_size_wrapped(
            label, body_w, label_h, start_pt=_LABEL_START_PT, min_pt=_MIN_PT
        )
        label_box = add_textbox(
            slide,
            body_x,
            label_y,
            body_w,
            label_h,
            text=label,
            font_size_pt=label_pt,
            color=colors["muted"],
            word_wrap=True,
        )
        shapes.append(label_box)

    return shape_elements(shapes)
