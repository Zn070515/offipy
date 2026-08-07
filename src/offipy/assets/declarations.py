"""offipy.assets — HTML asset declaration parser and canonicalizer.

A3 Task 2 (rev1.2 §4.1/§4.2): reads canonical `data-asset`, legacy `data-icon`
and `data-primitive` sugar declarations from HTML, assigns deterministic
internal ids, and rewrites only the injected copy (user source HTML is never
edited in place).
"""

from __future__ import annotations

from dataclasses import dataclass
from html.parser import HTMLParser

from offipy.assets.model import (
    _PLACEMENTS,
    AssetPlacement,
    AssetRequest,
    canonical_params,
)
from offipy.assets.uri import parse_asset_uri
from offipy.exceptions import InvalidArgumentError

_LEGACY_ICON_SETS = frozenset({"ph", "lu"})
_PARAM_PREFIX = "data-asset-param-"


@dataclass(frozen=True)
class AssetDeclaration:
    declaration_id: str
    slide_index: int
    request: AssetRequest
    placement: AssetPlacement
    html_tag: str


def _line_starts(text: str) -> list[int]:
    starts = [0]
    for i, ch in enumerate(text):
        if ch == "\n":
            starts.append(i + 1)
    return starts


class _DeclarationExtractor(HTMLParser):
    def __init__(self, text: str) -> None:
        super().__init__(convert_charrefs=False)
        self._line_starts = _line_starts(text)
        self.declarations: list[AssetDeclaration] = []
        self._insertions: list[tuple[int, str]] = []
        self._slide_opened = 0
        self._slide_depth = 0
        self._ordinals: dict[int, int] = {}

    def _abs_pos(self) -> int:
        lineno, offset = self.getpos()
        return self._line_starts[lineno - 1] + offset

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        normalized: list[tuple[str, str]] = [(k, v or "") for k, v in attrs]
        d = dict(normalized)
        if "data-offipy-asset-id" in d:
            raise InvalidArgumentError("user source HTML must not carry data-offipy-asset-id")
        if tag == "section" and "data-pptx-slide" in d:
            self._slide_opened += 1
            self._slide_depth += 1
        has_asset = "data-asset" in d
        has_icon = "data-icon" in d
        has_prim = "data-primitive" in d
        count = int(has_asset) + int(has_icon) + int(has_prim)
        if count == 0:
            return
        if self._slide_depth == 0:
            raise InvalidArgumentError(
                f"asset declaration outside a <section data-pptx-slide> (element <{tag}>)"
            )
        if count > 1:
            raise InvalidArgumentError(
                f"element <{tag}> declares more than one of data-asset/data-icon/data-primitive"
            )
        start_abs = self._abs_pos()
        raw = self.get_starttag_text()
        assert raw is not None  # get_starttag_text is set right after a start tag
        end_abs = start_abs + len(raw)
        insert_at = end_abs - (2 if raw.endswith("/>") else 1)

        slide_index = self._slide_opened
        if has_asset:
            uri = d["data-asset"]
        elif has_icon:
            uri = self._legacy_icon_uri(d["data-icon"])
        else:
            uri = self._primitive_uri(d["data-primitive"])
        request = self._build_request(uri, normalized, slide_index)
        placement = self._read_placement(d, slide_index)
        ordinal = self._ordinals.get(slide_index, 0) + 1
        self._ordinals[slide_index] = ordinal
        declaration_id = f"asset-s{slide_index:02d}-{ordinal:03d}"

        self.declarations.append(
            AssetDeclaration(
                declaration_id=declaration_id,
                slide_index=slide_index,
                request=request,
                placement=placement,
                html_tag=tag,
            )
        )
        new_attrs: list[str] = []
        if not has_asset:
            new_attrs.append(f'data-asset="{uri}"')
        if "data-asset-placement" not in d:
            new_attrs.append(f'data-asset-placement="{placement}"')
        new_attrs.append(f'data-offipy-asset-id="{declaration_id}"')
        self._insertions.append((insert_at, " " + " ".join(new_attrs)))

    def handle_endtag(self, tag: str) -> None:
        if tag == "section" and self._slide_depth > 0:
            self._slide_depth -= 1

    @staticmethod
    def _legacy_icon_uri(data_icon: str) -> str:
        set_, sep, name = data_icon.partition(":")
        if not sep or set_ not in _LEGACY_ICON_SETS or not name:
            raise InvalidArgumentError(
                f"invalid legacy data-icon {data_icon!r} (expected ph:<name> or lu:<name>)"
            )
        return f"asset://{set_}/icon/{name}"

    @staticmethod
    def _primitive_uri(name: str) -> str:
        if not name:
            raise InvalidArgumentError("data-primitive name must be non-empty")
        return f"asset://primitives/primitive/{name}"

    def _build_request(
        self, uri: str, attrs: list[tuple[str, str]], slide_index: int
    ) -> AssetRequest:
        base = parse_asset_uri(uri)
        attr_params = [
            (k[len(_PARAM_PREFIX) :], v) for k, v in attrs if k.startswith(_PARAM_PREFIX)
        ]
        if not attr_params:
            return base
        merged = list(base.params) + attr_params
        try:
            params = canonical_params(merged)
        except InvalidArgumentError as exc:
            raise InvalidArgumentError(f"slide {slide_index} asset declaration: {exc}") from exc
        return AssetRequest(base.ref, params)

    @staticmethod
    def _read_placement(d: dict[str, str], slide_index: int) -> AssetPlacement:
        placement = d.get("data-asset-placement", "") or "replace"
        if placement not in _PLACEMENTS:
            raise InvalidArgumentError(
                f"slide {slide_index} asset declaration has invalid "
                f"data-asset-placement {placement!r}"
            )
        return placement  # type: ignore[return-value]


def preprocess_asset_declarations(html_text: str) -> tuple[str, list[AssetDeclaration]]:
    """Parse asset declarations and return (rewritten HTML, declarations).

    The rewritten HTML carries canonical `data-asset`, resolved
    `data-asset-placement` and a deterministic `data-offipy-asset-id` on every
    declared element; all other bytes (text, comments, other attributes) are
    preserved verbatim. Returns the input unchanged when nothing is declared.
    """
    extractor = _DeclarationExtractor(html_text)
    extractor.feed(html_text)
    extractor.close()
    if not extractor._insertions:
        return html_text, extractor.declarations
    out = html_text
    for pos, s in sorted(extractor._insertions, key=lambda t: t[0], reverse=True):
        out = out[:pos] + s + out[pos:]
    return out, extractor.declarations
