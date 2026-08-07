"""offipy.assets.primitives.browser_mockup — editable browser mockup (A5).

An outer window card (``fill``, default surface) with a chrome bar, three
native window dots, an optional centered title, a small accent highlight, an
optional URL pill (text only, no hyperlink side effect), and an empty white
content viewport (no screenshot in v0.14). Title wraps to fit; the URL is
single-line and raises if it cannot fit at the floor.
"""

from __future__ import annotations

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

_CHROME = "#D9DDE3"
_VIEWPORT = "#FFFFFF"
_WHITE = "#FFFFFF"
_DOT_COLORS = ("#FF5F57", "#FEBC2E", "#28C840")
_CHROME_H_FRACTION = 0.12
_URL_H_FRACTION = 0.10
_PAD_FRACTION = 0.03
_MIN_PT = 7.0


def render(slide, params, context) -> tuple[object, ...]:
    """Draw the browser mockup and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    pad = max(_PAD_FRACTION * min(w, h), 3.0)
    chrome_h = h * _CHROME_H_FRACTION

    shapes = []
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

    chrome = add_shape(
        slide,
        MSO_SHAPE.RECTANGLE,
        x,
        y,
        w,
        chrome_h,
        fill=_CHROME,
        line="transparent",
    )
    shapes.append(chrome)

    d = chrome_h * 0.5
    dot_y = y + (chrome_h - d) / 2
    for i, color in enumerate(_DOT_COLORS):
        dot = add_shape(
            slide,
            MSO_SHAPE.OVAL,
            x + pad + i * d * 1.4,
            dot_y,
            d,
            d,
            fill=color,
            line="transparent",
        )
        shapes.append(dot)

    accent_x = x + pad + 3 * d * 1.4 + pad
    accent = add_shape(
        slide,
        MSO_SHAPE.RECTANGLE,
        accent_x,
        y + chrome_h - 3,
        0.12 * w,
        3,
        fill=colors["accent"],
        line="transparent",
    )
    shapes.append(accent)

    title = params.get("title")
    if title:
        title_pt = fit_font_size_wrapped(
            title, w - 2 * pad, chrome_h, start_pt=chrome_h * 0.45, min_pt=_MIN_PT
        )
        title_box = add_textbox(
            slide,
            x + pad,
            y,
            w - 2 * pad,
            chrome_h,
            text=title,
            font_size_pt=title_pt,
            color=colors["muted"],
            word_wrap=True,
        )
        shapes.append(title_box)

    viewport_top = y + chrome_h + pad
    url = params.get("url")
    if url:
        url_h = h * _URL_H_FRACTION
        url_pill = add_shape(
            slide,
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x + pad,
            viewport_top,
            w - 2 * pad,
            url_h,
            fill=_WHITE,
            line=colors["accent"],
        )
        shapes.append(url_pill)
        url_pt = fit_font_size(url, w - 2 * pad, url_h, start_pt=url_h * 0.5, min_pt=_MIN_PT)
        url_box = add_textbox(
            slide,
            x + pad,
            viewport_top,
            w - 2 * pad,
            url_h,
            text=url,
            font_size_pt=url_pt,
            color=colors["muted"],
        )
        shapes.append(url_box)
        viewport_top = viewport_top + url_h + pad

    viewport_h = y + h - pad - viewport_top
    viewport = add_shape(
        slide,
        MSO_SHAPE.RECTANGLE,
        x + pad,
        viewport_top,
        w - 2 * pad,
        viewport_h,
        fill=_VIEWPORT,
        line="transparent",
    )
    shapes.append(viewport)

    return shape_elements(shapes)
