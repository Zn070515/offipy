"""offipy.assets — canonical asset URI parsing and formatting.

Contract (rev1.2 §3.7): `asset://<provider>/<kind>/<name>[?k=v&...]`. Query keys
are canonicalized (`_` -> `-`, lower-cased, sorted) and must be unique after
canonicalization; values are percent-decoded. `#` fragments and CSS `var(`
references are rejected. Color hex values round-trip as `%23RRGGBB`.
"""

from __future__ import annotations

import urllib.parse
from typing import cast

from offipy.assets.model import AssetKind, AssetRef, AssetRequest, canonical_params
from offipy.exceptions import InvalidArgumentError

_SCHEME = "asset://"
_HEX = frozenset("0123456789abcdefABCDEF")
_UNRESERVED = frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")


def _percent_decode(value: str) -> str:
    if "%" not in value:
        return value
    i = 0
    n = len(value)
    while i < n:
        if value[i] == "%":
            hexpart = value[i + 1 : i + 3]
            if len(hexpart) != 2 or hexpart[0] not in _HEX or hexpart[1] not in _HEX:
                raise InvalidArgumentError("malformed percent escape in asset URI")
            i += 3
        else:
            i += 1
    try:
        return urllib.parse.unquote_to_bytes(value).decode("utf-8")
    except UnicodeDecodeError as exc:
        raise InvalidArgumentError("asset URI percent-encoded value is not valid UTF-8") from exc


def _percent_encode(value: str) -> str:
    out: list[str] = []
    for ch in value:
        if ch in _UNRESERVED:
            out.append(ch)
        else:
            out.extend(f"%{b:02X}" for b in ch.encode("utf-8"))
    return "".join(out)


def _parse_query(query: str) -> tuple[tuple[str, str], ...]:
    pairs: list[tuple[str, str]] = []
    for pair in query.split("&"):
        if not pair:
            continue
        key, _sep, value = pair.partition("=")
        key = _percent_decode(key)
        value = _percent_decode(value)
        if "var(" in value:
            raise InvalidArgumentError("asset param value must not reference a CSS variable")
        pairs.append((key, value))
    return canonical_params(pairs)


def parse_asset_uri(uri: str) -> AssetRequest:
    """Parse an `asset://` URI into a canonical request; raises on any deviation."""
    if not isinstance(uri, str):
        raise InvalidArgumentError("asset URI must be a string")
    if not uri.startswith(_SCHEME):
        raise InvalidArgumentError(f"asset URI must start with {_SCHEME!r}")
    rest = uri[len(_SCHEME) :]
    if "#" in rest:
        raise InvalidArgumentError("asset URI must not contain a fragment")
    path_part, _sep, query_part = rest.partition("?")
    segments = path_part.split("/")
    if len(segments) != 3 or not all(segments):
        raise InvalidArgumentError("asset URI must have exactly provider/kind/name path segments")
    provider, kind, name = (_percent_decode(s) for s in segments)
    # kind is validated at runtime by AssetRef.__post_init__; cast narrows it
    # for static typing from the raw path segment str.
    params = _parse_query(query_part) if query_part else ()
    return AssetRequest(AssetRef(provider, cast("AssetKind", kind), name), params)


def format_asset_uri(request: AssetRequest) -> str:
    """Render a canonical asset URI with sorted, percent-encoded query params."""
    if not isinstance(request, AssetRequest):
        raise InvalidArgumentError("format_asset_uri expects an AssetRequest")
    ref = request.ref
    uri = f"{_SCHEME}{ref.provider}/{ref.kind}/{_percent_encode(ref.name)}"
    if request.params:
        parts: list[str] = []
        for key, value in request.params:
            if "var(" in value:
                raise InvalidArgumentError("asset param value must not reference a CSS variable")
            parts.append(f"{key}={_percent_encode(value)}")
        uri += "?" + "&".join(parts)
    return uri
