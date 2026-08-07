"""offipy.assets — public asset core (A2).

Pure stdlib import surface: model/uri/registry/license/color contracts live
here. Providers (icons, procedural, primitives) are added in A3/A4/A5.
"""

from offipy.assets.model import (
    AssetKind,
    AssetMeta,
    AssetPlacement,
    AssetProviderMeta,
    AssetRect,
    AssetRef,
    AssetRenderContext,
    AssetRequest,
    NativeShapePayload,
    RasterPayload,
    ResolvedAsset,
    SvgPayload,
    SvgTemplatePayload,
)

__all__ = [
    "AssetKind",
    "AssetMeta",
    "AssetPlacement",
    "AssetProviderMeta",
    "AssetRect",
    "AssetRef",
    "AssetRenderContext",
    "AssetRequest",
    "NativeShapePayload",
    "RasterPayload",
    "ResolvedAsset",
    "SvgPayload",
    "SvgTemplatePayload",
]
