"""offipy.assets — asset color value syntax validation.

Contract (rev1.2 §3.10): a color value is either a semantic theme token or a
strict hex `#RRGGBB` / `#RGB` (normalized uppercase). Tokens are not resolved
to theme colors here; that is deferred to theme materialization. CSS `var()`
references are rejected at the URI layer.
"""

from __future__ import annotations

import re

from offipy.exceptions import InvalidArgumentError

_TOKEN_ALIASES = {
    "accent": "accent",
    "accent-soft": "accent-soft",
    "ink": "ink",
    "muted": "muted",
    "surface": "surface",
    "background": "background",
    "bg": "background",
    "transparent": "transparent",
}

_HEX_COLOR_RE = re.compile(r"^#(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$")


def validate_color_value(value: str) -> str:
    """Validate and canonicalize a color value (token or hex), or raise."""
    if value in _TOKEN_ALIASES:
        return _TOKEN_ALIASES[value]
    if _HEX_COLOR_RE.match(value):
        return "#" + value[1:].upper()
    raise InvalidArgumentError(f"invalid color value {value!r}")
