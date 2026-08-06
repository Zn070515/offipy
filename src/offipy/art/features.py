"""客观特征计算（几何 + 角色/颜色/字体 + 密度）。

rev2.1：_union_area 为真·区间并集；spacing 先行列聚类再算组内 gap；
visual_mass 的 dominant 取 ink 贡献最大者。
"""

from __future__ import annotations

from dataclasses import dataclass

from .models import ArtElement, ArtSlide

_SKIP_ROLES = {
    "background",
    "container",
    "decoration",
    "page_number",
    "footer",
    "section",
    "closing",
}


@dataclass(frozen=True)
class AlignmentLine:
    position: float
    members: tuple[str, ...]
    axis: str  # "vertical" | "horizontal"

    def to_dict(self) -> dict:
        return {"position": self.position, "members": list(self.members), "axis": self.axis}


def _cluster_items(values: list[float], ids: list[str], tol: float = 0.01) -> list[list[str]]:
    """把一维坐标聚成对齐簇：相邻差值 ≤ tol 归一组。"""
    pairs = sorted(zip(values, ids, strict=False))
    clusters: list[tuple[float, list[str]]] = []
    for v, i in pairs:
        if clusters and abs(v - clusters[-1][0]) <= tol:
            clusters[-1][1].append(i)
        else:
            clusters.append((v, [i]))
    return [members for _v, members in clusters if len(members) >= 3]


def alignment_features(elements: list[ArtElement]) -> dict:
    """共享边对齐线：left/center/right + top/middle/bottom，每条 ≥3 成员。"""
    lines: list[AlignmentLine] = []
    els = [e for e in elements if e.area > 0]
    for axis, attrs in (("vertical", ("x", "cx", "right")), ("horizontal", ("y", "cy", "bottom"))):
        for name in attrs:
            vals = []
            ids = []
            for e in els:
                if name == "x":
                    v = e.x
                elif name == "cx":
                    v = e.x + e.width / 2
                elif name == "right":
                    v = e.x + e.width
                elif name == "y":
                    v = e.y
                elif name == "cy":
                    v = e.y + e.height / 2
                else:  # bottom
                    v = e.y + e.height
                vals.append(v)
                ids.append(e.element_id)
            for members in _cluster_items(vals, ids):
                pos = vals[ids.index(members[0])]
                lines.append(AlignmentLine(position=pos, members=tuple(members), axis=axis))
    return {"lines": lines}


def _gaps(elements: list[ArtElement], axis: str) -> list[float]:
    """组内相邻间隙：按 center 排序，间隙 = span - 两个半宽；重叠时 clamp 到 0。"""
    if len(elements) < 2:
        return []

    def center(e):
        return (e.x + e.width / 2) if axis == "x" else (e.y + e.height / 2)

    def half(e):
        return (e.width / 2) if axis == "x" else (e.height / 2)

    ordered = sorted(elements, key=center)
    out: list[float] = []
    for cur, nxt in zip(ordered, ordered[1:], strict=False):
        span = center(nxt) - center(cur)
        gap = span - half(cur) - half(nxt)
        out.append(round(max(gap, 0.0), 6))
    return out


def _center(e: ArtElement, axis: str) -> float:
    return (e.x + e.width / 2) if axis == "x" else (e.y + e.height / 2)


def _row_clusters(elements: list[ArtElement], tol: float = 0.02) -> list[list[ArtElement]]:
    """按 center-y 聚成行：相邻行中心距 ≤ tol 归一组。"""
    ordered = sorted(elements, key=lambda e: _center(e, "y"))
    rows: list[list[ArtElement]] = []
    for e in ordered:
        if rows and abs(_center(e, "y") - _center(rows[-1][0], "y")) <= tol:
            rows[-1].append(e)
        else:
            rows.append([e])
    return rows


def _column_clusters(elements: list[ArtElement], tol: float = 0.02) -> list[list[ArtElement]]:
    """按 center-x 聚成列。"""
    ordered = sorted(elements, key=lambda e: _center(e, "x"))
    cols: list[list[ArtElement]] = []
    for e in ordered:
        if cols and abs(_center(e, "x") - _center(cols[-1][0], "x")) <= tol:
            cols[-1].append(e)
        else:
            cols.append([e])
    return cols


def _drifted_count(values: list[float], tol: float | None = None) -> int:
    """偏离中位数的元素数。容差 = max(0.5×|中位数|, 0.01)。"""
    if len(values) < 3:
        return 0
    ordered = sorted(values)
    median = ordered[len(ordered) // 2]
    t = tol if tol is not None else max(0.5 * abs(median), 0.01)
    return sum(1 for v in values if abs(v - median) > t)


def _one_axis_spacing(elements: list[ArtElement], axis: str) -> dict:
    """行/列聚类后，组内算 gap 漂移；跨组不产生相邻。"""
    groups = _row_clusters(elements) if axis == "x" else _column_clusters(elements)
    gaps: list[float] = []
    for g in groups:
        gaps.extend(_gaps(g, axis))
    if len(gaps) < 2:
        return {"gaps": gaps, "drift_count": 0, "max_drift_ratio": 0.0}
    ordered = sorted(gaps)
    median = ordered[len(ordered) // 2]
    if median == 0:
        return {"gaps": gaps, "drift_count": 0, "max_drift_ratio": 0.0}
    drifts = [g - median for g in gaps]
    drift_count = sum(1 for d in drifts if abs(d) > max(0.5 * median, 0.01))
    return {
        "gaps": gaps,
        "drift_count": drift_count,
        "max_drift_ratio": max(abs(d) for d in drifts) / median,
    }


def spacing_features(elements: list[ArtElement]) -> dict:
    """水平/垂直组内间隙漂移统计（先行列聚类，跨组不误报）。"""
    return {
        "horizontal": _one_axis_spacing(elements, "x"),
        "vertical": _one_axis_spacing(elements, "y"),
    }


def _element_ink(e: ArtElement) -> float:
    """单元素 ink 贡献（与 visual_mass 的公式一致）。"""
    if e.kind == "image":
        return e.area * 2.0
    if e.kind == "text":
        size = e.font_size_norm or 0.01
        return max(len(e.text), 1) * size
    return e.area * 0.5


def visual_mass(elements: list[ArtElement]) -> dict:
    """视觉重量：按角色过滤后的 ink 质量；dominant 取 ink 贡献最大者。"""
    ink = 0.0
    count = 0
    dominant = None
    dominant_ink = -1.0
    for e in elements:
        if e.role in _SKIP_ROLES or e.is_background:
            continue
        count += 1
        contrib = _element_ink(e)
        ink += contrib
        if contrib > dominant_ink:
            dominant_ink = contrib
            dominant = e
    return {
        "ink": round(ink, 6),
        "count": count,
        "dominant_id": dominant.element_id if dominant else None,
    }


def _union_area(elements: list[ArtElement]) -> float:
    """真·区间并集面积：按所有 x 边界切段，每段做 y interval union。

    两个矩形垂直方向中间有空白 → 该空白不算入覆盖面积。
    """
    if not elements:
        return 0.0
    xs = sorted({e.x for e in elements} | {e.x + e.width for e in elements})
    if len(xs) < 2:
        return 0.0
    total = 0.0
    for x0, x1 in zip(xs, xs[1:], strict=False):
        if x1 - x0 <= 0:
            continue
        # 覆盖该 x 段的元素，收集其 y 区间
        intervals = sorted(
            (e.y, e.y + e.height) for e in elements if e.x <= x0 + (x1 - x0) / 2 <= e.x + e.width
        )
        # interval union：合并重叠区间后求和
        merged: list[tuple[float, float]] = []
        for lo, hi in intervals:
            if merged and lo <= merged[-1][1]:
                prev = merged[-1]
                merged[-1] = (prev[0], max(prev[1], hi))
            else:
                merged.append((lo, hi))
        covered_y = sum(hi - lo for lo, hi in merged)
        total += covered_y * (x1 - x0)
    return total


def density_features(elements: list[ArtElement]) -> dict:
    els = [e for e in elements if e.role not in _SKIP_ROLES and not e.is_background]
    sum_area = sum(e.area for e in els)
    union = _union_area(els)
    return {"sum_area_ratio": round(sum_area, 6), "union_area_ratio": round(union, 6)}


def physical_aspect_ratio(el: ArtElement, page_width: float, page_height: float) -> float:
    """把归一化宽高还原成物理尺寸后算宽高比。"""
    pw = el.width * page_width
    ph = el.height * page_height
    return (pw / ph) if ph else 0.0


def compute_features(slide: ArtSlide) -> dict:
    """角色/颜色/字体/焦点综合特征入口。

    Task 5 实现（当前为占位，避免导入期缺失 name 破坏 Task 4 测试集合）。
    """
    raise NotImplementedError
