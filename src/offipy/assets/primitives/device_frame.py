"""offipy.assets.primitives.device_frame — editable device frames (A5).

Phone/tablet/desktop bodies drawn entirely from native shapes with an empty
screen (no screenshot in v0.14). The design is scaled to preserve a sensible
device aspect and centered inside the rect, leaving unused space transparent.
A rect that scales the device below a frozen minimum raises rather than
producing degenerate geometry.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.shapes import MSO_SHAPE

from offipy.assets.primitives._common import (
    add_shape,
    require_rect,
    resolve_native_colors,
    shape_elements,
)
from offipy.exceptions import InvalidArgumentError

if TYPE_CHECKING:
    from collections.abc import Mapping

    from offipy.assets.model import AssetRect, AssetRenderContext

_BEZEL = "#1F2937"
_ASPECT = {"phone": 0.5, "tablet": 0.75, "desktop": 1.5}
_MIN_W_PX = 16.0
_MIN_H_PX = 24.0


def _fit(rect: AssetRect, aspect: float) -> tuple[float, float, float, float]:
    """Center a device of the given width/height aspect inside the rect."""
    x, y, w, h = rect.x, rect.y, rect.width, rect.height
    if w / h >= aspect:
        dev_w, dev_h = h * aspect, h
    else:
        dev_w, dev_h = w, w / aspect
    return x + (w - dev_w) / 2, y + (h - dev_h) / 2, dev_w, dev_h


def _require_device_size(dev_w: float, dev_h: float) -> None:
    if dev_w < _MIN_W_PX or dev_h < _MIN_H_PX:
        raise InvalidArgumentError(
            f"device rect too small for a legible frame ({dev_w:.0f}x{dev_h:.0f}px, "
            f"minimum {_MIN_W_PX:.0f}x{_MIN_H_PX:.0f}px)"
        )


def _phone(slide: Any, colors: Mapping[str, str], rect: AssetRect) -> list[Any]:
    x, y, w, h = _fit(rect, _ASPECT["phone"])
    _require_device_size(w, h)
    t = max(0.045 * w, 2.0)
    shapes = []
    body = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        x,
        y,
        w,
        h,
        fill=_BEZEL,
        line="transparent",
    )
    shapes.append(body)
    screen_x, screen_y = x + t, y + t * 1.5
    screen_w, screen_h = w - 2 * t, h - 2 * t * 1.5
    screen = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        screen_x,
        screen_y,
        screen_w,
        screen_h,
        fill=colors["fill"],
        line="transparent",
    )
    shapes.append(screen)
    notch_w, notch_h = 0.35 * screen_w, 0.035 * screen_h
    notch = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        screen_x + (screen_w - notch_w) / 2,
        screen_y,
        notch_w,
        notch_h,
        fill=_BEZEL,
        line="transparent",
    )
    shapes.append(notch)
    home_w, home_h = 0.28 * screen_w, 0.03 * screen_h
    home = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        screen_x + (screen_w - home_w) / 2,
        screen_y + screen_h - home_h,
        home_w,
        home_h,
        fill=colors["accent"],
        line="transparent",
    )
    shapes.append(home)
    return shapes


def _tablet(slide: Any, colors: Mapping[str, str], rect: AssetRect) -> list[Any]:
    x, y, w, h = _fit(rect, _ASPECT["tablet"])
    _require_device_size(w, h)
    t = max(0.045 * w, 2.0)
    shapes = []
    body = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        x,
        y,
        w,
        h,
        fill=_BEZEL,
        line="transparent",
    )
    shapes.append(body)
    screen_x, screen_y = x + t, y + t
    screen_w, screen_h = w - 2 * t, h - 2 * t
    screen = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        screen_x,
        screen_y,
        screen_w,
        screen_h,
        fill=colors["fill"],
        line="transparent",
    )
    shapes.append(screen)
    dot_d = 0.03 * screen_w
    dot = add_shape(
        slide,
        MSO_SHAPE.OVAL,
        screen_x + (screen_w - dot_d) / 2,
        screen_y + dot_d * 0.5,
        dot_d,
        dot_d,
        fill=_BEZEL,
        line="transparent",
    )
    shapes.append(dot)
    return shapes


def _desktop(slide: Any, colors: Mapping[str, str], rect: AssetRect) -> list[Any]:
    x, y, w, h = _fit(rect, _ASPECT["desktop"])
    _require_device_size(w, h)
    t = max(0.03 * w, 2.0)
    monitor_h = 0.78 * h
    neck_w = 0.12 * w
    neck_h = 0.12 * h
    base_h = 0.10 * h
    shapes = []
    monitor = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        x,
        y,
        w,
        monitor_h,
        fill=_BEZEL,
        line="transparent",
    )
    shapes.append(monitor)
    screen = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        x + t,
        y + t,
        w - 2 * t,
        monitor_h - 2 * t,
        fill=colors["fill"],
        line="transparent",
    )
    shapes.append(screen)
    neck = add_shape(
        slide,
        MSO_SHAPE.RECTANGLE,
        x + (w - neck_w) / 2,
        y + monitor_h,
        neck_w,
        neck_h,
        fill=_BEZEL,
        line="transparent",
    )
    shapes.append(neck)
    base = add_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        x + 0.08 * w,
        y + monitor_h + neck_h,
        0.84 * w,
        base_h,
        fill=colors["accent"],
        line="transparent",
    )
    shapes.append(base)
    return shapes


def render(
    slide: Any,
    params: Mapping[str, str],
    context: AssetRenderContext,
) -> tuple[object, ...]:
    """Draw the device frame and return its XML elements bottom → top."""
    rect = require_rect(context)
    colors = resolve_native_colors(params, context)
    device = params["device"]
    if device == "phone":
        shapes = _phone(slide, colors, rect)
    elif device == "tablet":
        shapes = _tablet(slide, colors, rect)
    else:
        shapes = _desktop(slide, colors, rect)
    return shape_elements(shapes)
