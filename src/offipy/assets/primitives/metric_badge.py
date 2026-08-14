"""offipy.assets.primitives.metric_badge — editable metric badge (A5).

Vertical stack: optional delta (top-right, muted), the value (largest, ink,
bold), then an optional label (muted). A rounded surface card sits behind it
when ``fill`` is not transparent. Delta is text only in v1 — no semantic
positive/negative color is inferred from ``+/-``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

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

_PAD_FRACTION = 0.06
_VALUE_H_FRACTION = 0.45
_DELTA_H_FRACTION = 0.18
_MIN_PT = 8.0
_MIN_PT_SMALL = 6.0


def render(
    slide: Any,
    params: Mapping[str, str],
    context: AssetRenderContext,
) -> tuple[object, ...]:
    """Draw the metric badge and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    pad = max(_PAD_FRACTION * min(w, h), 4.0)

    shapes = []
    if colors["fill"] != "transparent":
        card = add_shape(
            slide,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x,
            y,
            w,
            h,
            fill=colors["fill"],
            line="transparent",
        )
        shapes.append(card)

    inner_x = x + pad
    inner_w = w - 2 * pad
    inner_h = h - 2 * pad
    inner_top = y + pad

    delta = params.get("delta")
    label = params.get("label")
    delta_h = inner_h * _DELTA_H_FRACTION if delta else 0.0
    if label:
        value_h = inner_h * _VALUE_H_FRACTION if delta else inner_h * 0.5
        label_h = inner_h - delta_h - value_h
    else:
        value_h = inner_h - delta_h
        label_h = 0.0

    if delta:
        delta_pt = fit_font_size(
            delta,
            inner_w,
            delta_h,
            start_pt=min(delta_h * 0.7, 20.0),
            min_pt=_MIN_PT_SMALL,
        )
        delta_box = add_textbox(
            slide,
            inner_x,
            inner_top,
            inner_w,
            delta_h,
            text=delta,
            font_size_pt=delta_pt,
            color=colors["muted"],
            align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        shapes.append(delta_box)

    value_pt = fit_font_size_wrapped(
        params["value"], inner_w, value_h, start_pt=value_h * 0.85, min_pt=_MIN_PT
    )
    value_box = add_textbox(
        slide,
        inner_x,
        inner_top + delta_h,
        inner_w,
        value_h,
        text=params["value"],
        font_size_pt=value_pt,
        color=colors["ink"],
        bold=True,
        word_wrap=True,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    shapes.append(value_box)

    if label:
        label_pt = fit_font_size_wrapped(
            label, inner_w, label_h, start_pt=label_h * 0.6, min_pt=_MIN_PT_SMALL
        )
        label_box = add_textbox(
            slide,
            inner_x,
            inner_top + delta_h + value_h,
            inner_w,
            label_h,
            text=label,
            font_size_pt=label_pt,
            color=colors["muted"],
            word_wrap=True,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        shapes.append(label_box)

    return shape_elements(shapes)
