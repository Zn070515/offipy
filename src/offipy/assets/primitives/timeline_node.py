"""offipy.assets.primitives.timeline_node — phase-aware editable node (A5).

A single node marker (one primitive, no connecting lines) with an optional
editable label. Phase changes deterministic styling:
- past: muted solid marker + muted label;
- current: accent solid marker + ink label;
- future: muted outline marker + muted label.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.shapes import MSO_SHAPE

from offipy.assets.primitives._common import (
    add_shape,
    add_textbox,
    fit_font_size_wrapped,
    require_rect,
    resolve_native_colors,
    shape_elements,
)

if TYPE_CHECKING:
    from collections.abc import Mapping

    from offipy.assets.model import AssetRenderContext

_MARKER_D_FRACTION = 0.6
_PAD_FRACTION = 0.05
_MIN_PT = 8.0


def render(
    slide: Any,
    params: Mapping[str, str],
    context: AssetRenderContext,
) -> tuple[object, ...]:
    """Draw the timeline node and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    pad = max(_PAD_FRACTION * min(w, h), 4.0)

    d = h * _MARKER_D_FRACTION
    phase = params.get("phase", "current")
    marker_x = x + pad
    marker_y = y + (h - d) / 2

    if phase == "current":
        # fill 公共参数驱动 marker 填充；未传（transparent 默认）回退 accent（#59）
        marker_fill = colors["fill"] if colors["fill"] != "transparent" else colors["accent"]
        marker = add_shape(
            slide,
            MSO_SHAPE.OVAL,
            marker_x,
            marker_y,
            d,
            d,
            fill=marker_fill,
            line="transparent",
        )
    elif phase == "past":
        marker = add_shape(
            slide,
            MSO_SHAPE.OVAL,
            marker_x,
            marker_y,
            d,
            d,
            fill=colors["muted"],
            line="transparent",
        )
    else:  # future: muted outline
        marker = add_shape(
            slide,
            MSO_SHAPE.OVAL,
            marker_x,
            marker_y,
            d,
            d,
            fill="transparent",
            line=colors["muted"],
        )
    shapes = [marker]

    label = params.get("label")
    if label:
        body_x = x + pad + d + pad
        body_w = w - pad - d - 2 * pad
        label_pt = fit_font_size_wrapped(
            label, body_w, h - 2 * pad, start_pt=h * 0.4, min_pt=_MIN_PT
        )
        label_color = colors["muted"] if phase != "current" else colors["ink"]
        label_box = add_textbox(
            slide,
            body_x,
            y + pad,
            body_w,
            h - 2 * pad,
            text=label,
            font_size_pt=label_pt,
            color=label_color,
            word_wrap=True,
        )
        shapes.append(label_box)

    return shape_elements(shapes)
