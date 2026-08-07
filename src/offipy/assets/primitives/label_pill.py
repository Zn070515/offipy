"""offipy.assets.primitives.label_pill — editable label pill (A5).

Frozen style: rounded rectangle filled with the resolved ``fill`` common
param (falls back to the accent token / ``accent`` param when fill is unset),
with a single line of contrast text centered both axes. Text color is white
on dark fills and ink on light fills. A text too long for the rect shrinks;
if it cannot fit even at the 8pt floor the renderer raises rather than
enlarging the pill or truncating the text.
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from offipy.assets.primitives._common import (
    add_shape,
    add_textbox,
    fit_font_size,
    require_rect,
    resolve_native_colors,
    shape_elements,
)

_MIN_PT = 8.0
_LUMINANCE_THRESHOLD = 140.0


def _contrast_text(hex_color: str) -> str:
    """Return white for dark fills and ink for light fills."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return "#FFFFFF" if lum < _LUMINANCE_THRESHOLD else "#222222"


def render(slide, params, context) -> tuple[object, ...]:
    """Draw the label pill and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    accent = colors["accent"]
    # fill 公共参数驱动卡片填充；未传（transparent 是默认）时回退 accent（#59）
    pill_fill = colors["fill"] if colors["fill"] != "transparent" else accent

    pill = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        x,
        y,
        w,
        h,
        fill=pill_fill,
        line="transparent",
    )
    text_pt = fit_font_size(params["text"], w, h, start_pt=h * 0.6, min_pt=_MIN_PT)
    label = add_textbox(
        slide,
        x,
        y,
        w,
        h,
        text=params["text"],
        font_size_pt=text_pt,
        color=_contrast_text(pill_fill),
        bold=True,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    return shape_elements([pill, label])
