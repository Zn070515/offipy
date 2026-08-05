"""PPTX 审计规则：注册表驱动，不手工硬编码调用。

AuditRule Protocol：`run(context) -> list[AuditFinding]`。DEFAULT_RULES 依次执行
（Bounds 先于 Margin，写 context.bounds_edges 供 Margin 避免同方向双报）。

角色分类先于规则执行（roles.classify_presentation）。用户 ignore 与角色豁免的
margin 由中央 suppression 处理——全部进 suppressed 带 reason，不静默丢弃；
卡片容器（文本位于填充 AutoShape 内）由 OverlapRule 直接记 intentional_containment。

共同豁免：hidden（不可见无意义）、geometry_unknown（无法精确定位）。
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Protocol

from .extract import _ShapeRecord
from .geometry import Rect, overlap_area, rect_contains, rect_intersection
from .models import (
    AuditConfig,
    AuditFinding,
    AuditShapeRef,
    Severity,
    SuppressedFinding,
    SuppressionReason,
)
from .roles import classify_presentation

_TINY_AREA = 0.0025  # in²，极小装饰点
_MIN_DIM = 1e-6  # in，退化矩形（水平/垂直线条）判空宽高阈值
_FULL_COVER_RATIO = 0.98
_PARTIAL_RATIO = 0.5
_OFF_CANVAS_MID_AREA = 1.0  # in²
_HIGH_OVERSHOOT_FRACTION = 0.25
_EDGE_RULE = {
    "left": "geometry.margin.left",
    "right": "geometry.margin.right",
    "top": "geometry.margin.top",
    "bottom": "geometry.margin.bottom",
}
_EDGE_CN = {"left": "左", "right": "右", "top": "上", "bottom": "下"}


# ---------------------------------------------------------------- 上下文


@dataclass
class RuleContext:
    slide_size: tuple[float, float]
    config: AuditConfig
    records: list[_ShapeRecord]
    # BoundsRule 填写，MarginRule 读取：形状 (slide, id) → 已报越界的方向
    bounds_edges: dict[tuple[int, int], set[str]] = field(default_factory=dict)
    suppressed: list[SuppressedFinding] = field(default_factory=list)

    @property
    def slide_w(self) -> float:
        return self.slide_size[0]

    @property
    def slide_h(self) -> float:
        return self.slide_size[1]

    @property
    def page_rect(self) -> Rect:
        return Rect(0, 0, self.slide_w, self.slide_h)


class AuditRule(Protocol):
    def run(self, context: RuleContext) -> list[AuditFinding]: ...


# ---------------------------------------------------------------- 工具


def _shape_ref(rec: _ShapeRecord) -> AuditShapeRef:
    return AuditShapeRef(
        slide_index=rec.slide_index,
        shape_id=rec.shape_id,
        name=rec.name,
        shape_type=rec.shape_type,
        role=rec.role,
    )


def _rect(rec: _ShapeRecord) -> Rect | None:
    if rec.left is None or rec.top is None or rec.width is None or rec.height is None:
        return None
    return Rect(rec.left, rec.top, rec.width, rec.height)


def _finding(
    rule_id: str,
    kind: str,
    severity: Severity,
    message: str,
    primary: AuditShapeRef,
    secondary: AuditShapeRef | None = None,
    details: dict | None = None,
    confidence: float = 1.0,
) -> AuditFinding:
    return AuditFinding(
        rule_id=rule_id,
        kind=kind,  # type: ignore[arg-type]
        severity=severity,
        message=message,
        primary=primary,
        secondary=secondary,
        details=details or {},
        confidence=confidence,
    )


def _find_record(
    records: list[_ShapeRecord], slide_index: int, shape_id: int
) -> _ShapeRecord | None:
    for rec in records:
        if rec.slide_index == slide_index and rec.shape_id == shape_id:
            return rec
    return None


def _is_picture_chart(rec: _ShapeRecord) -> bool:
    return rec.shape_type in ("PICTURE", "CHART") or rec.has_table


def _same_parent(a: _ShapeRecord, b: _ShapeRecord) -> bool:
    return a.parent_shape_id == b.parent_shape_id


# ---------------------------------------------------------------- Bounds


class BoundsRule:
    rule_id = "bounds"

    def run(self, context: RuleContext) -> list[AuditFinding]:
        findings: list[AuditFinding] = []
        page = context.page_rect
        tol = context.config.bounds_tolerance_in
        for rec in context.records:
            if rec.is_group or rec.is_hidden or rec.geometry_unknown:
                continue
            r = _rect(rec)
            if r is None:
                continue
            if r.width < _MIN_DIM or r.height < _MIN_DIM:
                continue  # 退化矩形（水平/垂直线条），无面积可言，越界判定无意义
            over = {
                "left": -r.x,
                "right": r.right - page.right,
                "top": -r.y,
                "bottom": r.bottom - page.bottom,
            }
            off_edges = {e for e, d in over.items() if d > tol}
            if off_edges:
                context.bounds_edges[(rec.slide_index, rec.shape_id)] = off_edges
            if rect_intersection(r, page) is None:
                sev = Severity.MID if r.area() >= _OFF_CANVAS_MID_AREA else Severity.LOW
                findings.append(
                    _finding(
                        rule_id="geometry.bounds.off_canvas",
                        kind="bounds",
                        severity=sev,
                        message=(
                            f"形状完全在幻灯片画布外（{_off_canvas_desc(r, page)}），"
                            "可能是暂存/动画/设计残留"
                        ),
                        primary=_shape_ref(rec),
                        details={
                            "page_w_in": round(page.width, 4),
                            "page_h_in": round(page.height, 4),
                            "shape_w_in": round(r.width, 4),
                            "shape_h_in": round(r.height, 4),
                        },
                    )
                )
                continue
            if not off_edges:
                continue
            inter = rect_intersection(r, page)
            out_ratio = 1.0 - (inter.area() / r.area()) if inter is not None else 1.0
            max_over = max(over.values())
            high = out_ratio > 0.5 or max_over > _HIGH_OVERSHOOT_FRACTION * max(
                context.slide_w, context.slide_h
            )
            details = {
                "edges": sorted(off_edges),
                "max_overshoot_in": round(max_over, 4),
                "out_ratio": round(out_ratio, 4),
                "page_w_in": round(page.width, 4),
                "page_h_in": round(page.height, 4),
            }
            if high:
                findings.append(
                    _finding(
                        rule_id="geometry.bounds.partial",
                        kind="bounds",
                        severity=Severity.HIGH,
                        message=(
                            f"形状大面积越出幻灯片：{_edge_cn(off_edges)} 边，"
                            f"最大超出 {max_over:.3f} 英寸"
                            f"（页面 {page.width:.3f}×{page.height:.3f} 英寸）"
                        ),
                        primary=_shape_ref(rec),
                        details=details,
                    )
                )
            else:
                findings.append(
                    _finding(
                        rule_id="geometry.bounds.partial",
                        kind="bounds",
                        severity=Severity.MID,
                        message=(
                            f"形状部分越出幻灯片边界：{_edge_cn(off_edges)} 边，"
                            f"超出 {max_over:.3f} 英寸"
                            f"（页面 {page.width:.3f}×{page.height:.3f} 英寸）"
                        ),
                        primary=_shape_ref(rec),
                        details=details,
                    )
                )
        return findings


def _edge_cn(edges: set[str]) -> str:
    return "/".join(_EDGE_CN[e] for e in sorted(edges))


def _off_canvas_desc(r: Rect, page: Rect) -> str:
    parts = []
    if r.right <= 0:
        parts.append("整块在页面左侧")
    if r.x >= page.right:
        parts.append("整块在页面右侧")
    if r.bottom <= 0:
        parts.append("整块在页面上方")
    if r.y >= page.bottom:
        parts.append("整块在页面下方")
    return "/".join(parts) if parts else "与画布无交集"


# ---------------------------------------------------------------- Margin


class MarginRule:
    rule_id = "margin"

    def run(self, context: RuleContext) -> list[AuditFinding]:
        findings: list[AuditFinding] = []
        page = context.page_rect
        safe = context.config.safe_margin_in
        for rec in context.records:
            if rec.is_group or rec.is_hidden or rec.is_connector or rec.geometry_unknown:
                continue
            r = _rect(rec)
            if r is None:
                continue
            bounded = context.bounds_edges.get((rec.slide_index, rec.shape_id), set())
            gaps = {
                "left": r.x,
                "right": page.right - r.right,
                "top": r.y,
                "bottom": page.bottom - r.bottom,
            }
            for edge, gap in gaps.items():
                if edge in bounded:
                    continue  # bounds 已报该方向，不双报
                if 0.0 <= gap < safe:
                    findings.append(
                        _finding(
                            rule_id=_EDGE_RULE[edge],
                            kind="margin",
                            severity=Severity.LOW,
                            message=(
                                f"内容贴近幻灯片{_EDGE_CN[edge]}边缘："
                                f"实际间距 {gap:.3f} 英寸，要求 ≥ {safe} 英寸"
                            ),
                            primary=_shape_ref(rec),
                            details={
                                "edge": edge,
                                "gap_in": round(gap, 4),
                                "required_in": safe,
                                "page_w_in": round(page.width, 4),
                                "page_h_in": round(page.height, 4),
                            },
                        )
                    )
        return findings


# ---------------------------------------------------------------- Overlap


class OverlapRule:
    rule_id = "overlap"

    def run(self, context: RuleContext) -> list[AuditFinding]:
        findings: list[AuditFinding] = []
        by_slide: dict[int, list[_ShapeRecord]] = defaultdict(list)
        for rec in context.records:
            if rec.is_connector or rec.is_hidden or rec.is_group or rec.geometry_unknown:
                continue
            if rec.role == "background":
                continue  # 全页背景不参与普通 overlap
            r = _rect(rec)
            if r is None or r.area() < _TINY_AREA:
                continue
            by_slide[rec.slide_index].append(rec)
        for recs in by_slide.values():
            for i in range(len(recs)):
                for j in range(i + 1, len(recs)):
                    a, b = recs[i], recs[j]
                    if _related(a, b):
                        continue
                    f = _classify_overlap(a, b, context)
                    if f is not None:
                        findings.append(f)
        return findings


def _related(a: _ShapeRecord, b: _ShapeRecord) -> bool:
    """父子 / 祖孙对不枚举。"""
    return (
        a.parent_shape_id == b.shape_id
        or b.parent_shape_id == a.shape_id
        or a.shape_id in b.group_path
        or b.shape_id in a.group_path
    )


def _classify_overlap(
    a: _ShapeRecord, b: _ShapeRecord, context: RuleContext
) -> AuditFinding | None:
    ra, rb = _rect(a), _rect(b)
    assert ra is not None and rb is not None
    inter = rect_intersection(ra, rb)
    if inter is None:
        return None
    min_area = min(ra.area(), rb.area())
    if min_area <= 0:
        return None
    ratio = overlap_area(ra, rb) / min_area
    if ratio < _PARTIAL_RATIO:
        return None

    confidence = 0.5 if (a.is_rotated or b.is_rotated) else 1.0
    approx_note = "（旋转包围盒近似）" if confidence < 1.0 else ""
    same_parent = _same_parent(a, b)
    area_a, area_b = ra.area(), rb.area()

    if rect_contains(ra, rb, eps=1e-6) or rect_contains(rb, ra, eps=1e-6):
        return _contained_result(a, b, ratio, context, same_parent, confidence, approx_note)
    if ratio >= _FULL_COVER_RATIO:
        on_top, below = _on_top_below(a, b, area_a, area_b, same_parent)
        high = _is_picture_chart(on_top) and bool(below.text.strip())
        return _cover_finding(below, on_top, ratio, context, confidence, approx_note, high=high)
    if same_parent and a.parent_shape_id is not None:
        return _partial_finding(a, b, ratio, context, Severity.LOW, confidence, approx_note)
    return _partial_finding(a, b, ratio, context, Severity.MID, confidence, approx_note)


def _on_top_below(
    a: _ShapeRecord,
    b: _ShapeRecord,
    area_a: float,
    area_b: float,
    same_parent: bool,
) -> tuple[_ShapeRecord, _ShapeRecord]:
    """谁在上层：同容器内按 z-order（后加在上）；跨容器用面积大者兜底。"""
    if same_parent:
        return (a, b) if a.z_order >= b.z_order else (b, a)
    return (a, b) if area_a >= area_b else (b, a)


def _contained_result(
    a: _ShapeRecord,
    b: _ShapeRecord,
    ratio: float,
    context: RuleContext,
    same_parent: bool,
    confidence: float,
    approx_note: str,
) -> AuditFinding | None:
    """一方完全包含另一方（含相等矩形）时的分类。on_top 决定文本是否被遮挡。"""
    ra, rb = _rect(a), _rect(b)
    assert ra is not None and rb is not None
    on_top, below = _on_top_below(a, b, ra.area(), rb.area(), same_parent)
    if _is_picture_chart(on_top):
        # 图片/图表盖住：下方有文本 → 内容被遮挡 HIGH；否则图片盖图片 MID
        return _cover_finding(
            below, on_top, ratio, context, confidence, approx_note, high=bool(below.text.strip())
        )
    if on_top.has_text_frame and on_top.text.strip():
        if below.shape_type == "AUTO_SHAPE" and same_parent:
            # 文本位于填充 AutoShape 内 → 卡片容器，不报 overlap
            context.suppressed.append(
                SuppressedFinding(
                    finding=_cover_finding(on_top, below, ratio, context, 1.0, ""),
                    reason="intentional_containment",
                )
            )
            return None
        if _is_picture_chart(below):
            return None  # 文本题注盖在图上 → 正常配图
        return _cover_finding(below, on_top, ratio, context, confidence, approx_note)
    if below.has_text_frame and below.text.strip():
        # 文本在下层被非图片容器盖住 → 内容被遮挡
        return _cover_finding(below, on_top, ratio, context, confidence, approx_note)
    return _cover_finding(below, on_top, ratio, context, confidence, approx_note)


def _overlap_area(a: _ShapeRecord, b: _ShapeRecord) -> float:
    ra, rb = _rect(a), _rect(b)
    if ra is None or rb is None:
        return 0.0
    return overlap_area(ra, rb)


def _cover_finding(
    covered: _ShapeRecord,
    cover: _ShapeRecord,
    ratio: float,
    context: RuleContext,
    confidence: float,
    approx_note: str,
    high: bool = False,
) -> AuditFinding:
    if high:
        severity = Severity.HIGH
        message = (
            f"文本被图片/图表完全覆盖，内容可能被遮挡：形状 #{covered.shape_id} "
            f"完全位于 #{cover.shape_id} 之下（重叠比 {ratio:.2f}）{approx_note}"
        )
    else:
        severity = Severity.MID
        message = (
            f"形状完全覆盖另一个形状：覆盖比 {ratio:.2f}，"
            f"#{covered.shape_id} 被 #{cover.shape_id} 完全覆盖{approx_note}"
        )
    return _finding(
        rule_id="geometry.overlap.covered_text",
        kind="overlap",
        severity=severity,
        message=message,
        primary=_shape_ref(covered),
        secondary=_shape_ref(cover),
        details={
            "overlap_ratio": round(ratio, 4),
            "overlap_area_in2": round(_overlap_area(covered, cover), 4),
            "approx": confidence < 1.0,
        },
        confidence=confidence,
    )


def _partial_finding(
    a: _ShapeRecord,
    b: _ShapeRecord,
    ratio: float,
    context: RuleContext,
    severity: Severity,
    confidence: float,
    approx_note: str,
) -> AuditFinding:
    primary = a if a.shape_id <= b.shape_id else b
    secondary = b if primary is a else a
    message = (
        f"形状部分重叠：覆盖比 {ratio:.2f}（以较小形状计），"
        f"#{primary.shape_id} 与 #{secondary.shape_id}{approx_note}"
    )
    return _finding(
        rule_id="geometry.overlap.partial",
        kind="overlap",
        severity=severity,
        message=message,
        primary=_shape_ref(primary),
        secondary=_shape_ref(secondary),
        details={
            "overlap_ratio": round(ratio, 4),
            "overlap_area_in2": round(_overlap_area(a, b), 4),
            "approx": confidence < 1.0,
        },
        confidence=confidence,
    )


# ---------------------------------------------------------------- 编排


DEFAULT_RULES: list[AuditRule] = [BoundsRule(), MarginRule(), OverlapRule()]

_ROLE_MARGIN_REASON: dict[str, SuppressionReason] = {
    "background": "full_bleed",
    "page_number": "page_number",
    "header": "header_footer",
    "footer": "header_footer",
    "decoration": "repeated_decoration",
}


def run_rules(
    records: list[_ShapeRecord],
    slide_size: tuple[float, float],
    config: AuditConfig,
) -> tuple[list[AuditFinding], list[SuppressedFinding]]:
    """分类角色 → 逐规则执行 → 中央豁免进 suppressed。"""
    classify_presentation(records, slide_size)
    context = RuleContext(slide_size=slide_size, config=config, records=records)
    findings: list[AuditFinding] = []
    for rule in DEFAULT_RULES:
        for f in rule.run(context):
            reason = _suppression_reason(f, context)
            if reason is not None:
                context.suppressed.append(SuppressedFinding(finding=f, reason=reason))
            else:
                findings.append(f)
    return findings, context.suppressed


def _suppression_reason(f: AuditFinding, context: RuleContext) -> SuppressionReason | None:
    slide_index = f.primary.slide_index
    shape_id = f.primary.shape_id
    cfg = context.config
    if cfg.ignored_shapes and (slide_index, shape_id) in cfg.ignored_shapes:
        return "user_shape"
    rec = _find_record(context.records, slide_index, shape_id)
    if rec is not None and cfg.ignored_regions:
        r = _rect(rec)
        if r is not None:
            cx, cy = r.center()
            for rx, ry, rw, rh in cfg.ignored_regions:
                if rx <= cx <= rx + rw and ry <= cy <= ry + rh:
                    return "user_region"
    if rec is not None and f.kind == "margin":
        return _ROLE_MARGIN_REASON.get(rec.role)
    return None
