"""offipy.assets — deferred SVG theme materialization (rev1.2 §3.10/§3.11).

A2 freezes reusable color logic without rendering PPTX: `resolve_asset_color`
turns a token/hex/transparent value into a concrete color given theme vars,
and `materialize_svg_template` substitutes a template's color slots with
resolved colors, validates the result is well-formed XML, and returns an
immutable `SvgPayload` preserving the template view box.
"""

from __future__ import annotations

import re
from collections.abc import Mapping

from offipy.assets._xml import parse_svg
from offipy.assets.color import _HEX_COLOR_RE, validate_color_value
from offipy.assets.model import SvgPayload, SvgTemplatePayload
from offipy.exceptions import InvalidArgumentError

_SENTINEL_RE = re.compile(r"__[A-Za-z0-9_-]+__")


def _norm_theme_key(key: str) -> str:
    k = key.strip().lower().replace("_", "-")
    if k.startswith("--"):
        k = k[2:]
    return k


def _normalized_theme_vars(theme_vars: Mapping[str, str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for key, value in theme_vars.items():
        out.setdefault(_norm_theme_key(key), value)
    return out


def resolve_asset_color(value: str, theme_vars: Mapping[str, str]) -> str:
    """Resolve a token/hex/transparent color value against theme vars."""
    canon = validate_color_value(value)
    if canon == "transparent" or canon.startswith("#"):
        return canon
    resolved = _normalized_theme_vars(theme_vars).get(canon)
    if not resolved:
        # None 或空串都视为「token 未定义」：deck 侧 measure 对未定义 CSS 变量
        # getPropertyValue 返回空串而非 None，空值按 not defined 报错，避免
        # 误导用户去检查 hex 格式（#60）。
        raise InvalidArgumentError(f"theme color token {canon!r} not defined")
    if not isinstance(resolved, str):
        raise InvalidArgumentError(f"theme color {canon!r} value must be a string")
    if resolved == "transparent":
        return "transparent"
    if _HEX_COLOR_RE.match(resolved):
        return "#" + resolved[1:].upper()
    raise InvalidArgumentError(
        f"theme color {canon!r} value {resolved!r} must be #RRGGBB or transparent"
    )


def materialize_svg_template(
    payload: SvgTemplatePayload, theme_vars: Mapping[str, str]
) -> SvgPayload:
    """Substitute declared color slots with resolved colors and validate the result."""
    resolved: dict[str, str] = {}
    for placeholder, value in payload.color_slots:
        if placeholder in resolved:
            raise InvalidArgumentError(f"duplicate color slot {placeholder!r}")
        resolved[placeholder] = resolve_asset_color(value, theme_vars)
    svg = payload.template
    for placeholder in resolved:
        if placeholder not in svg:
            raise InvalidArgumentError(
                f"declared color slot {placeholder!r} not present in template"
            )
    # longest-first so a placeholder that prefixes another is not clobbered
    for placeholder, color in sorted(resolved.items(), key=lambda kv: len(kv[0]), reverse=True):
        svg = svg.replace(placeholder, color)
    leftover = _SENTINEL_RE.search(svg)
    if leftover is not None:
        raise InvalidArgumentError(f"undeclared template sentinel {leftover.group(0)!r} remains")
    parse_svg(svg)
    return SvgPayload(svg=svg, render_mode="svg", view_box=payload.view_box)
