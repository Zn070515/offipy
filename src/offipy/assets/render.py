"""offipy.assets — generic asset measurement binding and placeholder lookup.

A3 Task 6: bind converter `kind=asset` measurements to deterministic
declaration ids, and locate the transparent placeholder in an assembled deck.
Internal-facing helpers for the postprocess pipeline; not part of the public
asset core surface.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from offipy.assets.declarations import AssetDeclaration
from offipy.assets.model import AssetRect
from offipy.exceptions import InvalidArgumentError

PLACEHOLDER_PREFIX = "OFFIPY_ASSET::"


@dataclass(frozen=True)
class AssetMeasurement:
    """converter `kind=asset` 测量记录的内部视图（非公共 Asset 模型）。"""

    asset_id: str
    slide_index: int
    rect: AssetRect
    theme_vars: dict[str, str]
    html_tag: str
    color: str | None = None

    def __post_init__(self) -> None:
        # 冻结上下文的 theme_vars 与调用方可变 dict 解耦
        object.__setattr__(self, "theme_vars", dict(self.theme_vars))


def _load_slides_data(data) -> list[dict]:
    if isinstance(data, dict) and "slides" in data:
        return data["slides"]
    if isinstance(data, list):
        return data
    raise InvalidArgumentError("measurements.json 缺少 slides 数组")


def load_asset_measurements(path: str | Path) -> dict[str, AssetMeasurement]:
    """读取 measurements.json，收集 kind=asset 记录并按 assetId 建索引。

    严格校验（不静默丢弃）：assetId 非空、全 deck 内唯一、rect 必须为正。
    返回 dict[declaration_id → AssetMeasurement]。
    """
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    result: dict[str, AssetMeasurement] = {}
    for i, sdata in enumerate(_load_slides_data(data)):
        for rec in sdata.get("records", []):
            if rec.get("kind") != "asset":
                continue
            asset_id = rec.get("assetId") or ""
            if not asset_id:
                raise InvalidArgumentError(f"slide {i + 1} 的 asset 测量记录缺少 assetId")
            if asset_id in result:
                raise InvalidArgumentError(f"asset 测量记录 assetId 重复: {asset_id}")
            rect = rec.get("rect") or {}
            try:
                ar = AssetRect(
                    x=float(rect["x"]),
                    y=float(rect["y"]),
                    width=float(rect["w"]),
                    height=float(rect["h"]),
                )
            except (KeyError, TypeError, ValueError) as exc:
                raise InvalidArgumentError(f"asset {asset_id} rect 无效: {rect!r}") from exc
            ar.validate_render()  # 非正 rect → 可执行错误，不静默丢弃
            color = rec.get("color")
            result[asset_id] = AssetMeasurement(
                asset_id=asset_id,
                slide_index=i + 1,
                rect=ar,
                theme_vars={k: str(v) for k, v in (rec.get("themeVars") or {}).items()},
                html_tag=str(rec.get("tag") or ""),
                color=str(color) if color is not None else None,
            )
    return result


def find_asset_placeholder(slide, declaration_id: str):
    """在已装配 slide 里按 name 精确查找透明占位符。

    必须恰好一个；0 或 >1 都是绑定错误。不做 outerHTML / 空间匹配兜底。
    """
    name = f"{PLACEHOLDER_PREFIX}{declaration_id}"
    matches = [sp for sp in slide.shapes if sp.name == name]
    if len(matches) != 1:
        raise InvalidArgumentError(
            f"asset 占位符 {declaration_id} 应为 1 个，实际 {len(matches)} 个"
        )
    return matches[0]


def bind_asset_measurements(
    declarations: list[AssetDeclaration],
    measurements: dict[str, AssetMeasurement],
) -> dict[str, AssetMeasurement]:
    """把声明按 declaration_id 绑定到测量，校验整副 deck 一致性。

    Fail on: 声明无测量、测量无声明、slide 序号不匹配。返回 dict 保持声明顺序。
    """
    bound: dict[str, AssetMeasurement] = {}
    for decl in declarations:
        measurement = measurements.get(decl.declaration_id)
        if measurement is None:
            raise InvalidArgumentError(f"asset 声明 {decl.declaration_id} 缺少对应测量记录")
        if measurement.slide_index != decl.slide_index:
            raise InvalidArgumentError(
                f"asset 声明 {decl.declaration_id} slide 序号不匹配："
                f"声明 slide {decl.slide_index} vs 测量 slide {measurement.slide_index}"
            )
        bound[decl.declaration_id] = measurement
    for asset_id in measurements:
        if asset_id not in bound:
            raise InvalidArgumentError(f"asset 测量记录 {asset_id} 无对应声明")
    return bound
