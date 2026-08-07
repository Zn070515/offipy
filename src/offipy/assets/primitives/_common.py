"""offipy.assets.primitives._common — shared native-rendering helpers (A5).

Font/theme policy: v0.14 AssetRenderContext carries color vars only. Colors
resolve from measurement theme vars via the A2 `resolve_asset_color` helper;
primitive labels use a stable common sans font (Arial), matching the deck
theme's `--font-sans`. No new public font token model is introduced.

Geometry discipline: every primitive draws inside its `AssetRect` (EMU via the
same 6350 px→EMU conversion as the converter) and returns python-pptx shapes /
XML elements in visual stacking order bottom → top. Renderers never touch
placeholder z-order themselves.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

from offipy.assets.materialize import resolve_asset_color
from offipy.assets.model import AssetRect, AssetRenderContext
from offipy.exceptions import InvalidArgumentError

PX_TO_EMU = 6350  # same conversion as the converter (1920×1080 px canvas)

# Stable common sans for primitive labels, matching deck theme --font-sans.
PRIMITIVE_FONT = "Arial"

_EMPTY_RECT_MIN = 1.0  # px


def px_to_emu(v: float) -> int:
    return int(round(v * PX_TO_EMU))


def require_rect(ctx: AssetRenderContext) -> AssetRect:
    """Validate the render rect is positive and return it."""
    ctx.rect.validate_render()
    return ctx.rect


def _resolve_text_token(ctx: AssetRenderContext, token: str, fallback: str) -> str:
    """Resolve a theme token for internal text color, falling back on absence.

    Text tokens (ink/muted) are rendering details, not user-facing params; a
    theme that omits them must not break a primitive.
    """
    try:
        return resolve_asset_color(token, ctx.theme_vars)
    except InvalidArgumentError:
        return fallback


def resolve_native_colors(params: Mapping[str, str], ctx: AssetRenderContext) -> dict[str, str]:
    """Resolve accent/fill (user-facing) plus ink/muted text tokens.

    accent/fill come from the validated payload params; unknown theme tokens
    there raise (matches A2). ink/muted are internal and fall back to safe
    defaults when the theme does not define them.
    """
    return {
        "accent": resolve_asset_color(params.get("accent", "accent"), ctx.theme_vars),
        "fill": resolve_asset_color(params.get("fill", "transparent"), ctx.theme_vars),
        "ink": _resolve_text_token(ctx, "ink", "#222222"),
        "muted": _resolve_text_token(ctx, "muted", "#667085"),
    }


def _rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.lstrip("#"))


def set_fill_line(shape, fill: str | None, line: str | None = None) -> None:
    """Set solid/transparent fill and line on a python-pptx shape.

    ``None`` or ``"transparent"`` means no fill (shape.fill.background()) /
    no line (line.fill.background()).
    """
    if fill is None or fill == "transparent":
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)
    if line is None or line == "transparent":
        shape.line.fill.background()
    else:
        shape.line.color.rgb = _rgb(line)


def add_shape(
    slide,
    shape_type: MSO_SHAPE,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str | None = None,
    line: str | None = None,
) -> Any:
    """Add a native autoshape at EMU coordinates, returning the python-pptx shape."""
    sp = slide.shapes.add_shape(shape_type, px_to_emu(x), px_to_emu(y), px_to_emu(w), px_to_emu(h))
    set_fill_line(sp, fill, line)
    return sp


def add_textbox(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    text: str = "",
    font_size_pt: float = 12.0,
    color: str = "#222222",
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
    word_wrap: bool = False,
) -> Any:
    """Add a textbox with a single editable run at EMU coordinates.

    Margins are zeroed so the box geometry matches the given rect exactly and
    text fills it deterministically.
    """
    tb = slide.shapes.add_textbox(px_to_emu(x), px_to_emu(y), px_to_emu(w), px_to_emu(h))
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = PRIMITIVE_FONT
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    return tb


def set_text_style(
    run,
    *,
    text: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
) -> None:
    """Re-style a text run in place (mutate an already-created run)."""
    if text is not None:
        run.text = text
    run.font.name = PRIMITIVE_FONT
    if font_size_pt is not None:
        run.font.size = Pt(font_size_pt)
    if color is not None:
        run.font.color.rgb = _rgb(color)
    if bold is not None:
        run.font.bold = bold


def shape_elements(shapes: Iterable) -> tuple[object, ...]:
    """Normalize python-pptx shapes / XML elements to their XML elements.

    Renderers return shapes; the generic placement code inserts their elements
    into the placeholder slot, so we must hand it XML elements (python-pptx
    shapes are wrapped; their ``_element`` is the actual OOXML).
    """
    return tuple(getattr(s, "_element", s) for s in shapes)


def fit_font_size(
    text: str,
    w_px: float,
    h_px: float,
    *,
    start_pt: float,
    min_pt: float,
    approx_char_width: float = 0.55,
    line_height_pt: float = 1.2,
) -> float:
    """Largest font size (pt) that fits a single line in w_px×h_px.

    Uses a rough proportional-width estimate (avg glyph width as a fraction of
    the em box). Never returns below ``min_pt``; if even ``min_pt`` overflows,
    raises instead of silently producing unreadable output.
    """
    if not text:
        return min(start_pt, max(min_pt, h_px * 72.0 / 96.0 / line_height_pt))
    for size_pt in _descending_pt(start_pt, min_pt):
        char_w_px = size_pt * (96.0 / 72.0) * approx_char_width
        line_h_px = size_pt * (96.0 / 72.0) * line_height_pt
        if char_w_px * len(text) <= w_px and line_h_px <= h_px:
            return size_pt
    raise InvalidArgumentError(
        f"text {text[:24]!r} does not fit rect {w_px:.0f}x{h_px:.0f}px at min {min_pt}pt"
    )


def fit_font_size_wrapped(
    text: str,
    w_px: float,
    h_px: float,
    *,
    start_pt: float,
    min_pt: float,
    approx_char_width: float = 0.55,
    line_height_pt: float = 1.2,
) -> float:
    """Largest font size (pt) whose greedy word-wrap fits w_px×h_px.

    Whitespace is the only wrap point; a token wider than the box must shrink
    (never truncates user text). Raises when even ``min_pt`` overflows either
    axis, so no rendered line ever escapes the rect.
    """
    if not text:
        return min(start_pt, max(min_pt, h_px * 72.0 / 96.0 / line_height_pt))
    for size_pt in _descending_pt(start_pt, min_pt):
        char_w_px = size_pt * (96.0 / 72.0) * approx_char_width
        line_h_px = size_pt * (96.0 / 72.0) * line_height_pt
        n_lines, max_line_w = _wrap_measure(text, char_w_px, w_px)
        if max_line_w <= w_px and n_lines * line_h_px <= h_px:
            return size_pt
    raise InvalidArgumentError(
        f"text {text[:24]!r} does not fit rect {w_px:.0f}x{h_px:.0f}px at min {min_pt}pt"
    )


def _wrap_measure(text: str, char_w_px: float, w_px: float) -> tuple[int, float]:
    space_w_px = char_w_px * 0.3
    lines = 1
    cur = 0.0
    max_line_w = 0.0
    for token in text.split():
        token_w = len(token) * char_w_px
        if cur == 0:
            cur = token_w
        elif cur + space_w_px + token_w <= w_px:
            cur += space_w_px + token_w
        else:
            max_line_w = max(max_line_w, cur)
            lines += 1
            cur = token_w
    return lines, max(max_line_w, cur)


def _descending_pt(start_pt: float, min_pt: float) -> Iterable[float]:
    step = 1.0
    size = start_pt
    while size >= min_pt - 1e-9:
        yield size
        size -= step
        if size < min_pt:
            step = 0.25
