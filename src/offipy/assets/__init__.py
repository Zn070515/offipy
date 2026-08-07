"""offipy.assets — public asset core (Asset System v1).

Pure stdlib import surface: model/uri/registry/license/color/materialize
contracts live here. Default providers: ph/lu (icons), procedural (patterns),
primitives (native presentation primitives).
"""

from offipy.assets.color import validate_color_value
from offipy.assets.license import ALLOWED_LICENSES, LicensePolicy
from offipy.assets.materialize import materialize_svg_template, resolve_asset_color
from offipy.assets.model import (
    AssetKind,
    AssetMeta,
    AssetPayload,
    AssetPlacement,
    AssetProvider,
    AssetProviderMeta,
    AssetRect,
    AssetRef,
    AssetRenderContext,
    AssetRenderMode,
    AssetRequest,
    NativeShapePayload,
    RasterPayload,
    ResolvedAsset,
    SvgPayload,
    SvgTemplatePayload,
)
from offipy.assets.registry import AssetRegistry, get_default_registry
from offipy.assets.uri import format_asset_uri, parse_asset_uri

__all__ = [
    "ALLOWED_LICENSES",
    "AssetKind",
    "AssetMeta",
    "AssetPayload",
    "AssetPlacement",
    "AssetProvider",
    "AssetProviderMeta",
    "AssetRect",
    "AssetRef",
    "AssetRegistry",
    "AssetRenderContext",
    "AssetRenderMode",
    "AssetRequest",
    "LicensePolicy",
    "NativeShapePayload",
    "RasterPayload",
    "ResolvedAsset",
    "SvgPayload",
    "SvgTemplatePayload",
    "format_asset_uri",
    "get_default_registry",
    "materialize_svg_template",
    "parse_asset_uri",
    "resolve_asset_color",
    "validate_color_value",
]
