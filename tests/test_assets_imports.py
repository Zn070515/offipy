"""A2 Task 6 — pure-stdlib import gate and public export surface."""

import os
import subprocess
import sys

import offipy.assets

_EXPECTED_EXPORTS = {
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
}


def test_public_exports_present():
    missing = _EXPECTED_EXPORTS - set(dir(offipy.assets))
    assert not missing, f"offipy.assets missing public exports: {sorted(missing)}"


def test_assets_import_without_heavy_deps():
    code = (
        "import sys\n"
        "import offipy\n"
        "import offipy.assets\n"
        "assert 'pptx' not in sys.modules\n"
        "assert 'PIL' not in sys.modules\n"
        "assert 'playwright' not in sys.modules\n"
    )
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, env=env)
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_reset_helper_is_not_public_export():
    assert "reset_default_registry_for_tests" not in set(dir(offipy.assets))
