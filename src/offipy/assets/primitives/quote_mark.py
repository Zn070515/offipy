"""offipy.assets.primitives.quote_mark — editable quote decoration (A5).

Composition: a large accent quotation glyph on the left, the user's text in an
editable wrapping textbox to its right, and an optional rounded card behind it
when ``fill`` is not transparent. Text always wraps inside the rect; at the
8pt floor an over-long quote raises rather than escaping the rect.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

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

_GLYPH = "“"  # left double quotation mark
_PAD_FRACTION = 0.06
_GLYPH_W_FRACTION = 0.22
_GLYPH_H_FRACTION = 0.30
_MIN_PT = 8.0
_TEXT_START_PT = 28.0


def render(
    slide: Any,
    params: Mapping[str, str],
    context: AssetRenderContext,
) -> tuple[object, ...]:
    """Draw the quote decoration and return its XML elements bottom → top."""
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

    gw = w * _GLYPH_W_FRACTION
    gh = h * _GLYPH_H_FRACTION
    glyph_pt = fit_font_size(_GLYPH, gw, gh, start_pt=gh * 0.7, min_pt=_MIN_PT)
    glyph = add_textbox(
        slide,
        x + pad,
        y + pad,
        gw,
        gh,
        text=_GLYPH,
        font_size_pt=glyph_pt,
        color=colors["accent"],
        bold=True,
        align=PP_ALIGN.LEFT,
    )
    shapes.append(glyph)

    text_x = x + pad + gw + pad
    text_w = w - pad - gw - 2 * pad
    text_pt = fit_font_size_wrapped(
        params["text"], text_w, h - 2 * pad, start_pt=_TEXT_START_PT, min_pt=_MIN_PT
    )
    textbox = add_textbox(
        slide,
        text_x,
        y + pad,
        text_w,
        h - 2 * pad,
        text=params["text"],
        font_size_pt=text_pt,
        color=colors["ink"],
        word_wrap=True,
    )
    shapes.append(textbox)

    return shape_elements(shapes)
