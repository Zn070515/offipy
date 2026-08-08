"""offipy.assets — safe XML parsing guard for vendored/user SVG content.

stdlib ``xml.etree.ElementTree`` expands internal general entities (billion-laughs
is possible) and would parse a DOCTYPE that smuggles entity bombs. Every SVG
parse in the assets pipeline goes through ``parse_svg``, which rejects any
DTD/entity declaration before the parser ever runs.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from offipy.exceptions import InvalidArgumentError

# Case-insensitive: the whole string is uppercased before matching, so
# `<!doctype` / `<!entity` are caught too. Any DTD/entity declaration is rejected.
_FORBIDDEN_MARKERS = ("<!DOCTYPE", "<!ENTITY", "<!ELEMENT", "<!ATTLIST", "<!NOTATION")


def parse_svg(svg: str) -> ET.Element:
    """Parse an SVG string, rejecting any DTD/entity declaration (XXE / billion-laughs guard)."""
    upper = svg.upper()
    for marker in _FORBIDDEN_MARKERS:
        if marker in upper:
            raise InvalidArgumentError("SVG must not declare a DOCTYPE or entity")
    try:
        return ET.fromstring(svg)
    except ET.ParseError as exc:
        raise InvalidArgumentError("SVG XML is malformed") from exc
