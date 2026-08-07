"""A5 Task 16/17 — v0.14 asset system release contract (CI-verifiable).

不需要构建 wheel：在源码树上直接验证「发布所依赖的公开资产契约」——
默认注册表四个 provider 端到端可解析、带必填参数的图元走 AssetRequest 不弱化
校验、八种 procedural / 八种 primitive 全部按 schema 校验后解析、非法参数在
provider 层拒绝。配合 test_assets_imports 的纯标准库 import 门禁共同构成
wheel/sdist 冒烟的可 CI 复现部分（Task 17 的真机 wheel 冒烟另走 pypi_smoke）。
"""

from __future__ import annotations

import pytest

from offipy.assets import (
    AssetRef,
    AssetRequest,
    NativeShapePayload,
    SvgPayload,
    get_default_registry,
)
from offipy.exceptions import InvalidArgumentError

_PH = 1512
_LU = 1756
_PROCEDURAL = 8
_PRIMITIVES = 8

# 必填参数走 AssetRequest，不弱化校验；timeline-node/browser-mockup 全参数可选
_PRIMITIVE_REQUIRED: dict[str, tuple[tuple[str, str], ...]] = {
    "quote-mark": (("text", "Work smarter"),),
    "section-number": (("number", "42"),),
    "label-pill": (("text", "Hot"),),
    "metric-badge": (("value", "24%"),),
    "process-arrow": (("steps", "Plan,Build"),),
    "device-frame": (("device", "phone"),),
}


def _resolve_by_ref(ref: AssetRef, params: tuple[tuple[str, str], ...] = ()) -> object:
    return get_default_registry().resolve(AssetRequest(ref, params))


# ---------------------------------------------------------------------------
# default registry: providers + search surface
# ---------------------------------------------------------------------------


class TestReleaseRegistry:
    def test_default_registry_provider_order_and_kinds(self) -> None:
        reg = get_default_registry()
        assert [reg.provider(p).provider_id for p in ("ph", "lu", "procedural", "primitives")] == [
            "ph",
            "lu",
            "procedural",
            "primitives",
        ]
        assert reg.provider("ph").kinds == frozenset({"icon"})
        assert reg.provider("lu").kinds == frozenset({"icon"})
        assert reg.provider("procedural").kinds == frozenset({"pattern"})
        assert reg.provider("primitives").kinds == frozenset({"primitive"})

    def test_search_surface_has_all_vendored_assets(self) -> None:
        metas = get_default_registry().search("", limit=10000)
        assert len(metas) == _PH + _LU + _PROCEDURAL + _PRIMITIVES

    def test_providers_report_first_party_and_license(self) -> None:
        reg = get_default_registry()
        assert reg.provider("primitives").provider_meta.first_party is True
        assert reg.provider("primitives").provider_meta.license == "MIT"
        assert reg.provider("procedural").provider_meta.first_party is True
        assert reg.provider("procedural").provider_meta.license == "MIT"


# ---------------------------------------------------------------------------
# end-to-end resolve per provider
# ---------------------------------------------------------------------------


class TestReleaseResolve:
    def test_icon_providers_resolve(self) -> None:
        ph = _resolve_by_ref(AssetRef("ph", "icon", "check"))
        assert isinstance(ph.payload, SvgPayload)
        assert ph.payload.render_mode == "freeform_svg"
        lu = _resolve_by_ref(AssetRef("lu", "icon", "settings"))
        assert isinstance(lu.payload, SvgPayload)

    @pytest.mark.parametrize(
        "pattern",
        [
            "wave",
            "blob",
            "dot-grid",
            "square-grid",
            "rings",
            "topography",
            "circuit",
            "gradient-orb",
        ],
    )
    def test_all_procedural_patterns_resolve(self, pattern: str) -> None:
        resolved = _resolve_by_ref(AssetRef("procedural", "pattern", pattern))
        assert resolved.payload.render_mode == "svg"

    @pytest.mark.parametrize(
        "primitive",
        [
            "quote-mark",
            "section-number",
            "label-pill",
            "metric-badge",
            "timeline-node",
            "process-arrow",
            "device-frame",
            "browser-mockup",
        ],
    )
    def test_all_primitives_resolve_with_required_params(self, primitive: str) -> None:
        params = _PRIMITIVE_REQUIRED.get(primitive, ())
        resolved = _resolve_by_ref(AssetRef("primitives", "primitive", primitive), params)
        assert isinstance(resolved.payload, NativeShapePayload)
        assert resolved.payload.primitive == primitive

    def test_optional_only_primitive_resolves_with_no_params(self) -> None:
        resolved = _resolve_by_ref(AssetRef("primitives", "primitive", "browser-mockup"))
        assert isinstance(resolved.payload, NativeShapePayload)

    def test_required_primitive_without_params_fails(self) -> None:
        with pytest.raises(InvalidArgumentError):
            _resolve_by_ref(AssetRef("primitives", "primitive", "label-pill"))

    def test_unknown_provider_or_kind_fails(self) -> None:
        with pytest.raises(InvalidArgumentError):
            _resolve_by_ref(AssetRef("bogus", "icon", "check"))
        with pytest.raises(InvalidArgumentError):
            _resolve_by_ref(AssetRef("ph", "pattern", "check"))

    def test_forbidden_screenshot_param_fails(self) -> None:
        with pytest.raises(InvalidArgumentError, match="screenshot"):
            _resolve_by_ref(
                AssetRef("primitives", "primitive", "device-frame"),
                (("device", "phone"), ("screenshot", "x")),
            )
