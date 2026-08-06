"""PPTX 审计规则：注册表驱动，不手工硬编码调用。

AuditRule Protocol：`run(context) -> list[AuditFinding]`。DEFAULT_RULES 依次执行
（Bounds 先于 Margin，写 context.bounds_edges 供 Margin 避免同方向双报）。

角色分类先于规则执行（roles.classify_presentation）。用户 ignore 与角色豁免的
margin 由中央 suppression 处理——全部进 suppressed 带 reason，不静默丢弃；
卡片容器（文本位于填充 AutoShape 内）由 OverlapRule 直接记 intentional_containment。

共同豁免：hidden（不可见无意义）、geometry_unknown（无法精确定位）。
"""

from __future__ import annotations

import functools
import math
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from .extract import _Paragraph, _ShapeRecord, _TextRun
from .geometry import Rect, overlap_area, rect_contains, rect_intersection
from .models import (
    RULE_AUTOFIT_GROW,
    RULE_AUTOFIT_SHRINK,
    RULE_TEXT_FIT_HORIZONTAL,
    RULE_TEXT_FIT_VERTICAL,
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
_DECORATIVE_SIZE_FRACTION = 0.5
# 装饰浮卡片豁免：on_top 无文本 + 面积 ≤ 卡片一半（≤ _DECORATIVE_SIZE_FRACTION ×
# 卡片面积）→ 小装饰/色条分层设计（实心不透明装饰仍按尺寸判别，透明装饰走
# transparent_overlay）。
_OFF_CANVAS_MID_AREA = 1.0  # in²
_HIGH_OVERSHOOT_FRACTION = 0.25
_EDGE_RULE = {
    "left": "geometry.margin.left",
    "right": "geometry.margin.right",
    "top": "geometry.margin.top",
    "bottom": "geometry.margin.bottom",
}
_EDGE_CN = {"left": "左", "right": "右", "top": "上", "bottom": "下"}

_MIN_READABLE_PT = 8.0  # 最小可读字号阈值
_DEFAULT_FONT_SIZE_PT = 18.0  # 文本框默认字号（未显式设置时的估算基准）
# 溢出噪声下限：Pillow(FreeType) 与 PowerPoint(DirectWrite) 度量 msyh 有亚 pt 级
# 引擎差。低于 1pt 的「超出」在 PowerPoint 里渲染刚好放下，报了是新的「脱节」。
_OVERFLOW_FLOOR_IN = 1.0 / 72.0
_LINE_HEIGHT_RATIO = 1.2
_PILLOW_CONF = 0.8  # Pillow 字体度量置信度
_FALLBACK_CONF = 0.4  # 字符权重回退置信度（消息标注「字符估算低置信」）
_ASCII_WEIGHT = 0.5  # 相对字号的 ASCII/数字宽度
_SPACE_WEIGHT = 0.35
_TEXT_FIT_SKIP_ROLES = ("page_number", "header", "footer")  # 页码/页眉页脚小文本本就紧凑


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
            if rec.role == "background" and context.config.ignore_full_bleed_shapes:
                continue  # 全页背景不参与普通 overlap（--no-full-bleed-ignore 可关闭）
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
    on_top, below = _on_top_below(a, b, area_a, area_b, same_parent)
    contained = rect_contains(ra, rb, eps=1e-6) or rect_contains(rb, ra, eps=1e-6)

    reason = _overlay_suppression_reason(on_top, below, contained)
    if reason is not None:
        if reason in ("text_on_background", "decorative_layering"):
            f = _partial_finding(a, b, ratio, context, Severity.LOW, confidence, approx_note)
        else:
            f = _cover_finding(below, on_top, ratio, context, confidence, approx_note)
        context.suppressed.append(SuppressedFinding(finding=f, reason=reason))
        return None

    if contained:
        return _contained_result(a, b, ratio, context, same_parent, confidence, approx_note)
    if ratio >= _FULL_COVER_RATIO:
        high = _is_picture_chart(on_top) and bool(below.text.strip())
        return _covered_or_none(below, on_top, ratio, context, confidence, approx_note, high=high)
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


def _overlay_suppression_reason(
    on_top: _ShapeRecord, below: _ShapeRecord, contained: bool
) -> SuppressionReason | None:
    """统一遮挡判定：返回豁免 reason，None 则按正常分类报。

    遮挡问题的本质是「上层是否真的盖住下层内容」：
    - 透明（a:noFill）无文本上层 → 视觉上不遮挡任何东西 → transparent_overlay；
    - 实心小装饰浮有文本容器（面积尺寸判别）→ decorative_overlay（包含/非包含都放行）；
    - 文字浮无文本背景/容器（非包含）→ text_on_background：下方无内容可遮挡，
      文字本身在上层完全可见，只是叠在背景条上。
    有文本的上层（透明与否）一律不豁免——文本叠文本是真问题，透明不改变内容冲突。
    """
    if _is_transparent(on_top) and not _has_text(on_top):
        return "transparent_overlay"
    if _is_decorative_overlay(on_top, below):
        return "decorative_overlay"
    if not contained and _has_text(on_top) and not _has_text(below):
        return "text_on_background"
    if not contained and _is_textless_decorative_layering(on_top, below):
        return "decorative_layering"
    return None


def _is_transparent(rec: _ShapeRecord) -> bool:
    """是否透明（a:noFill）——只认显式 noFill，继承填充按不透明处理。"""
    return rec.fill_kind == "none"


def _has_text(rec: _ShapeRecord) -> bool:
    return rec.has_text_frame and bool(rec.text.strip())


def _is_decorative_overlay(on_top: _ShapeRecord, below: _ShapeRecord) -> bool:
    """小装饰/色条浮在卡片内容器上（分层设计）→ 豁免 covered_text。

    on_top 无文本 + 面积 ≤ 卡片一半 + 双方非 PICTURE/CHART + 卡片有文本
    （有内容的容器）才算装饰；大块无文本形状盖住卡片（面积占优）不豁免；
    无文本纯背景不豁免（让背景在 --no-full-bleed-ignore 时参与 overlap）。
    """
    if _is_picture_chart(on_top) or _is_picture_chart(below):
        return False
    if _has_text(on_top):
        return False
    if not _has_text(below):
        return False
    r_on, r_below = _rect(on_top), _rect(below)
    if r_on is None or r_below is None:
        return False
    return r_on.area() <= r_below.area() * _DECORATIVE_SIZE_FRACTION


def _is_strip(r: Rect) -> bool:
    """长条：长边 ≥ 3 × 短边（宽卡/窄色条/行纹）。"""
    long_side = max(r.width, r.height)
    short_side = min(r.width, r.height)
    return short_side > 0 and long_side >= 3.0 * short_side


def _is_textless_decorative_layering(on_top: _ShapeRecord, below: _ShapeRecord) -> bool:
    """卡片容器 + 行条纹装饰分层（双方无文本）→ 豁免 partial。

    判别：双方无文本、非图片/图表、均为长条（aspect ≥ 3）、小者长边横跨 ≥ 70%
    大者长边（条纹铺满容器宽、仅边缘越出）。典型：S13 表格页卡片 Rectangle +
    行条纹 poking 出卡片底边——设计分层被报「部分重叠」。
    """
    if _is_picture_chart(on_top) or _is_picture_chart(below):
        return False
    if _has_text(on_top) or _has_text(below):
        return False
    r_on, r_below = _rect(on_top), _rect(below)
    if r_on is None or r_below is None:
        return False
    if not _is_strip(r_on) or not _is_strip(r_below):
        return False
    small, big = (r_on, r_below) if r_on.area() <= r_below.area() else (r_below, r_on)
    return max(small.width, small.height) >= 0.7 * max(big.width, big.height)


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
        # 图片/图表盖住：下方有文本 → 内容被遮挡 HIGH；下方无文本 → 无内容可遮挡
        return _covered_or_none(
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
        return _covered_or_none(below, on_top, ratio, context, confidence, approx_note)
    # 下方有文本 → 内容被遮挡；无文本 → 无内容可遮挡（装饰点浮空卡片不再误报）
    return _covered_or_none(below, on_top, ratio, context, confidence, approx_note)


def _overlap_area(a: _ShapeRecord, b: _ShapeRecord) -> float:
    ra, rb = _rect(a), _rect(b)
    if ra is None or rb is None:
        return 0.0
    return overlap_area(ra, rb)


def _covered_or_none(
    covered: _ShapeRecord,
    cover: _ShapeRecord,
    ratio: float,
    context: RuleContext,
    confidence: float,
    approx_note: str,
    high: bool = False,
) -> AuditFinding | None:
    """covered_text 只在被盖形状有文本时发射；下层无文本就无内容可遮挡。"""
    if not _has_text(covered):
        return None
    return _cover_finding(covered, cover, ratio, context, confidence, approx_note, high=high)


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


# ---------------------------------------------------------------- TextFit


def _is_wide_char(ch: str) -> bool:
    """东亚全宽字符近似（East Asian Width W/F）：CJK/全角/谚文等。"""
    code = ord(ch)
    return (
        0x1100 <= code <= 0x115F
        or 0x2E80 <= code <= 0xA4CF
        or 0xAC00 <= code <= 0xD7A3
        or 0xF900 <= code <= 0xFAFF
        or 0xFE30 <= code <= 0xFE4F
        or 0xFF00 <= code <= 0xFF60
        or 0xFFE0 <= code <= 0xFFE6
    )


def _char_width_pt(ch: str, size_pt: float) -> float:
    if ch == " ":
        return _SPACE_WEIGHT * size_pt
    return size_pt if _is_wide_char(ch) else _ASCII_WEIGHT * size_pt


# 已知字体 → (常规, 加粗) 文件。微软雅黑/宋体是 TTC 集合（msyh.ttc/simsun.ttc），
# 用 base 拼 .ttf 永远找不到 → 必须显式映射，否则 Pillow 度量全失败走字符权重回退。
_FONT_FILE_MAP: dict[str, tuple[str, str]] = {
    "microsoftyahei": ("msyh.ttc", "msyhbd.ttc"),
    "microsoftyaheiui": ("msyh.ttc", "msyhbd.ttc"),
    "yahei": ("msyh.ttc", "msyhbd.ttc"),
    "微软雅黑": ("msyh.ttc", "msyhbd.ttc"),
    "simsun": ("simsun.ttc", "simhei.ttf"),
    "宋体": ("simsun.ttc", "simhei.ttf"),
    "simhei": ("simhei.ttf", "simhei.ttf"),
    "黑体": ("simhei.ttf", "simhei.ttf"),
    "simkai": ("simkai.ttf", "simkai.ttf"),
    "kaiti": ("simkai.ttf", "simkai.ttf"),
    "kaitisc": ("simkai.ttf", "simkai.ttf"),
    "楷体": ("simkai.ttf", "simkai.ttf"),
    "simfang": ("simfang.ttf", "simfang.ttf"),
    "fangsong": ("simfang.ttf", "simfang.ttf"),
    "仿宋": ("simfang.ttf", "simfang.ttf"),
    "dengxian": ("deng.ttf", "dengbd.ttf"),
    "等线": ("deng.ttf", "dengbd.ttf"),
    "segoeui": ("segoeui.ttf", "segoeuib.ttf"),
    "segoeuilight": ("segoeuil.ttf", "segoeuil.ttf"),
    "calibri": ("calibri.ttf", "calibrib.ttf"),
    "arial": ("arial.ttf", "arialbd.ttf"),
    "timesnewroman": ("times.ttf", "timesbd.ttf"),
    "timesnewromanpsmt": ("times.ttf", "timesbd.ttf"),
    "cambria": ("cambria.ttc", "cambriab.ttf"),
    "cambriamath": ("cambria.ttc", "cambriab.ttf"),
    "consolas": ("consola.ttf", "consolab.ttf"),
    "couriernew": ("cour.ttf", "courbd.ttf"),
    "verdana": ("verdana.ttf", "verdanab.ttf"),
    "tahoma": ("tahoma.ttf", "tahomabd.ttf"),
    "georgia": ("georgia.ttf", "georgiab.ttf"),
    "lucidasansunicode": ("l_10646.ttf", "l_10646.ttf"),
    "lucidagrande": ("l_10646.ttf", "l_10646.ttf"),
    "msreferencesansserif": ("mssans.ttf", "mssansb.ttf"),
}


def _font_candidates(name: str, bold: bool) -> list[Path]:
    base = "".join(ch for ch in name if ch.isalnum())
    mapped = _FONT_FILE_MAP.get(base.lower())
    if mapped:
        # 已知字体：优先映射文件（加粗时先试加粗变体），再用 base 拼法兜底
        files = [mapped[1], mapped[0]] if bold else [mapped[0], mapped[1]]
        files.append(f"{base}.ttf")
    else:
        files = [f"{base}bd.ttf", f"{base}b.ttf"] if bold else []
        files.append(f"{base}.ttf")
    dirs = [
        Path(r"C:\Windows\Fonts"),
        Path("/usr/share/fonts/truetype/dejavu"),
        Path("/usr/share/fonts/truetype/msttcorefonts"),
        Path("/usr/share/fonts"),
        Path("/Library/Fonts"),
    ]
    cands = [d / f for d in dirs for f in files]
    if name.lower() in ("arial", "helvetica", "sans", "sans-serif", "calibri"):
        cands.append(Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
    return cands


def _load_font(font_name: str | None, bold: bool, size_pt: float):
    """定位字体文件加载 PIL 字体；失败返回 None（走字符权重回退）。"""
    try:
        from PIL import ImageFont
    except Exception:
        return None
    name = (font_name or "Arial").strip()
    for path in _font_candidates(name, bold):
        try:
            return ImageFont.truetype(str(path), int(size_pt))
        except Exception:
            continue
    try:
        return ImageFont.truetype(name, int(size_pt))  # 允许 font_name 本身是绝对路径
    except Exception:
        return None


def _pairpos_kerning(sub) -> dict[tuple[str, str], int]:
    """GPOS PairPos subtable → {(first_glyph, second_glyph): xadvance}。"""
    kern: dict[tuple[str, str], int] = {}
    if getattr(sub, "Format", None) == 1:
        firsts = sub.Coverage.glyphs
        for i, first in enumerate(firsts):
            for pv in sub.PairSet[i].PairValueRecord:
                v1 = pv.Value1
                xadv = v1.XAdvance if v1 and v1.XAdvance else 0
                if xadv:
                    kern[(first, pv.SecondGlyph)] = xadv
    elif getattr(sub, "Format", None) == 2:
        firsts = sub.Coverage.glyphs
        c1map = sub.ClassDef1.classDefs
        c2map = sub.ClassDef2.classDefs
        for first in firsts:
            c1 = c1map.get(first, 0)
            cr = sub.Class1Record[c1]
            for c2, rec in enumerate(cr.Class2Record):
                v1 = rec.Value1
                xadv = v1.XAdvance if v1 and v1.XAdvance else 0
                if not xadv:
                    continue
                for second, sc in c2map.items():
                    if sc == c2:
                        kern[(first, second)] = xadv
    return kern


def _font_kerning_pairs(tt) -> dict[tuple[str, str], int]:
    """{(left_glyph, right_glyph): xadvance font-units}。优先旧式 kern，无则 GPOS。"""
    kern: dict[tuple[str, str], int] = {}
    if "kern" in tt:
        for st in tt["kern"].kernTables:
            if st.format == 0 and (getattr(st, "coverage", 0) & 1):
                kern.update(st.kernTable)
    elif "GPOS" in tt:
        g = tt["GPOS"].table
        lookups = []
        for feat in g.FeatureList.FeatureRecord:
            if (feat.FeatureTag or "").lower() == "kern":
                lookups.extend(feat.Feature.LookupListIndex)
        if not lookups:
            lookups = list(range(len(g.LookupList.Lookup)))
        for idx in lookups:
            lookup = g.LookupList.Lookup[idx]
            if lookup.LookupType != 2:
                continue
            for sub in lookup.SubTable:
                kern.update(_pairpos_kerning(sub))
    return kern


@functools.lru_cache(maxsize=64)
def _font_metrics(
    path: str,
) -> tuple[float, dict[str, int], dict[tuple[str, str], int], dict[int, str]] | None:
    """(upem, {glyph: advance}, {(l, r): kern}, {codepoint: glyph})；失败 → None。"""
    try:
        from fontTools.ttLib import TTFont
    except Exception:
        return None
    try:
        tt = TTFont(path, fontNumber=0)
        upem = tt["head"].unitsPerEm
        advances = {gname: m[0] for gname, m in tt["hmtx"].metrics.items()}
        kern = _font_kerning_pairs(tt)
        cmap = tt.getBestCmap()
        return upem, advances, kern, cmap
    except Exception:
        return None


@functools.lru_cache(maxsize=32)
def _find_font_metrics(font_name: str, bold: bool):
    """定位字体文件并读 fontTools 度量；找不到/解析失败 → None。"""
    for path in _font_candidates(font_name, bold):
        if path.exists():
            m = _font_metrics(str(path))
            if m is not None:
                return m
    return _font_metrics(font_name)  # 允许 font_name 本身是绝对路径


def _run_width_pt(
    text: str, size_pt: float, bold: bool | None, font_name: str | None
) -> tuple[float, bool]:
    """单个 run 自然宽度（pt）→ (宽度, 是否高置信度量)。

    优先 fontTools hmtx+kerning：与 PowerPoint/DirectWrite 度量一致（拉丁/数字的
    kern 对如 T-e、R-T 会让 Pillow getlength 高估 ~1.3%）；无字体则 Pillow 度量；
    再失败则字符权重回退。
    """
    metrics = _find_font_metrics((font_name or "Arial").strip(), bool(bold))
    if metrics is not None:
        upem, advances, kern, cmap = metrics
        total = 0.0
        missing_pt = 0.0
        all_present = True
        names = []
        for ch in text:
            name = cmap.get(ord(ch))
            names.append(name)
            if name is None:
                # 缺码位（如 Arial 无 CJK 码位）绝不记 0 宽：按字符权重兜底
                # （全角 1em/半角 0.5em），并标记低置信走上层回退路径。
                all_present = False
                missing_pt += _char_width_pt(ch, size_pt)
            else:
                total += advances.get(name, 0)
        for i in range(len(names) - 1):
            gl, gr = names[i], names[i + 1]
            if gl is None or gr is None:
                continue
            total += kern.get((gl, gr), 0)
        return total / upem * size_pt + missing_pt, all_present
    font = _load_font(font_name, bool(bold), size_pt)
    if font is not None:
        try:
            return float(font.getlength(text)), True
        except Exception:
            pass
    return sum(_char_width_pt(ch, size_pt) for ch in text), False


def _visual_segments(p: _Paragraph) -> list[list[_TextRun]]:
    """a:br 拆出的视觉行，丢弃末尾空段（尾部 <a:br> 不渲染额外一行）。"""
    segs = p.segments if p.segments else [p.runs]
    while segs and not segs[-1]:
        segs = segs[:-1]
    return segs


def _para_segments_in(p: _Paragraph, default_size_pt: float) -> list[tuple[float, bool]]:
    """段落按 a:br 拆成视觉行，每行自然宽度（英寸）→ (宽, 是否用 Pillow 度量)。

    同一视觉行内的 run 求和；不同视觉行是独立行（wrap=none 时各自成行，
    wrap 时各自折行），绝不能跨行求和。
    """
    groups = _visual_segments(p)
    out = []
    for seg in groups:
        total_pt = 0.0
        used = False
        for run in seg:
            size = run.font_size or default_size_pt
            w_pt, u = _run_width_pt(run.text, size, run.bold, run.font_name)
            total_pt += w_pt
            used = used or u
        out.append((total_pt / 72.0, used))
    return out


def _para_size_pt(p: _Paragraph, default_size_pt: float) -> float:
    sizes = [r.font_size for r in p.runs if r.font_size is not None]
    return max(sizes) if sizes else default_size_pt


def _paragraph_line_height_pt(p: _Paragraph, size_pt: float) -> float:
    """段落单行行高（pt）：a:lnSpc 优先（spcPts 绝对 / spcPct 百分比），无则字号×1.2。"""
    if p.line_spacing_pts is not None:
        return p.line_spacing_pts
    if p.line_spacing_pct is not None:
        return size_pt * p.line_spacing_pct / 100.0
    return size_pt * _LINE_HEIGHT_RATIO


def _text_height_in(
    rec: _ShapeRecord, avail_w: float, default_size_pt: float
) -> tuple[float, bool]:
    """文本所需高（英寸，含折行）→ (高度, 是否用 Pillow 度量)。

    行高读段落 a:lnSpc（spcPts 绝对 / spcPct 百分比），未设则字号×1.2 兜底；
    wrap 时每段行数 = Σ max(1, ceil(段宽/可用宽))，no-wrap 时 = 视觉段数。
    """
    total_pt = 0.0
    used_pillow = False
    wraps = rec.word_wrap is not False  # None(未设) 按 PowerPoint 默认 square 折行
    for p in rec.paragraphs:
        segs_w = _para_segments_in(p, default_size_pt)
        used_pillow = used_pillow or any(u for _, u in segs_w)
        visual = _visual_segments(p)
        if wraps and avail_w > _MIN_DIM:
            lines = sum(max(1, math.ceil(w / avail_w)) for w, _ in segs_w)
        else:
            lines = len(visual) if visual else 1
        size = _para_size_pt(p, default_size_pt)
        total_pt += lines * _paragraph_line_height_pt(p, size)
    return total_pt / 72.0, used_pillow


def _effective_font_size_pt(rec: _ShapeRecord) -> float | None:
    sizes = [r.font_size for p in rec.paragraphs for r in p.runs if r.font_size is not None]
    return max(sizes) if sizes else None


def _text_overflow(
    rec: _ShapeRecord, avail_w: float, avail_h: float
) -> tuple[bool, bool, float, float]:
    """文本是否超出现有可用区域 → (超宽, 超高, 文本宽, 文本高)。"""
    segs_all = [_para_segments_in(p, _DEFAULT_FONT_SIZE_PT) for p in rec.paragraphs]
    max_w = max((w for segs in segs_all for w, _ in segs), default=0.0)
    wrap_w = avail_w if avail_w > _MIN_DIM else 1.0
    text_h, _ = _text_height_in(rec, wrap_w, _DEFAULT_FONT_SIZE_PT)
    # 仅显式 wrap=none 会横向溢出：square 自动折行，None(未设) 按 PowerPoint 默认 square
    over_w = (rec.word_wrap is False) and max_w > avail_w + _OVERFLOW_FLOOR_IN
    over_h = text_h > avail_h + _OVERFLOW_FLOOR_IN
    return over_w, over_h, max_w, text_h


class TextFitRule:
    rule_id = "text_fit"

    def run(self, context: RuleContext) -> list[AuditFinding]:
        findings: list[AuditFinding] = []
        for rec in context.records:
            if rec.is_group or rec.is_connector or rec.is_hidden:
                continue
            if (
                rec.has_table
                or rec.geometry_unknown
                or not rec.has_text_frame
                or not rec.text.strip()
            ):
                continue
            if rec.role in _TEXT_FIT_SKIP_ROLES and _role_text_fit_ignored(
                rec.role, context.config
            ):
                continue
            r = _rect(rec)
            if r is None:
                continue
            avail_w = r.width - (rec.tf_margin_left or 0.0) - (rec.tf_margin_right or 0.0)
            avail_h = r.height - (rec.tf_margin_top or 0.0) - (rec.tf_margin_bottom or 0.0)
            if avail_w <= _MIN_DIM or avail_h <= _MIN_DIM:
                findings.append(
                    _finding(
                        rule_id=RULE_TEXT_FIT_HORIZONTAL,
                        kind="text_fit",
                        severity=Severity.MID,
                        message=(
                            f"文本框无可用空间：可用宽 {max(avail_w, 0.0):.3f}×"
                            f"高 {max(avail_h, 0.0):.3f} 英寸"
                            f"（框 {r.width:.3f}×{r.height:.3f} 英寸，内边距吃尽）"
                        ),
                        primary=_shape_ref(rec),
                        details={
                            "avail_width_in": round(max(avail_w, 0.0), 4),
                            "avail_height_in": round(max(avail_h, 0.0), 4),
                        },
                        confidence=1.0,
                    )
                )
                continue
            segs_all = [_para_segments_in(p, _DEFAULT_FONT_SIZE_PT) for p in rec.paragraphs]
            max_para_w = max((w for segs in segs_all for w, _ in segs), default=0.0)
            used_pillow = any(u for segs in segs_all for _, u in segs)
            conf = _PILLOW_CONF if used_pillow else _FALLBACK_CONF
            approx = "" if used_pillow else "（字符估算低置信）"
            if rec.word_wrap is False and max_para_w > avail_w + _OVERFLOW_FLOOR_IN:
                findings.append(
                    _finding(
                        rule_id=RULE_TEXT_FIT_HORIZONTAL,
                        kind="text_fit",
                        severity=Severity.LOW,
                        message=(
                            f"文本横向超出文本框：自然宽 {max_para_w:.2f} 英寸 > "
                            f"可用 {avail_w:.2f} 英寸{approx}"
                        ),
                        primary=_shape_ref(rec),
                        details={
                            "text_width_in": round(max_para_w, 4),
                            "avail_width_in": round(avail_w, 4),
                            "word_wrap": bool(rec.word_wrap),
                        },
                        confidence=conf,
                    )
                )
            text_h, v_used = _text_height_in(rec, avail_w, _DEFAULT_FONT_SIZE_PT)
            v_conf = _PILLOW_CONF if (used_pillow or v_used) else _FALLBACK_CONF
            v_approx = "" if (used_pillow or v_used) else "（字符估算低置信）"
            if text_h > avail_h + _OVERFLOW_FLOOR_IN:
                findings.append(
                    _finding(
                        rule_id=RULE_TEXT_FIT_VERTICAL,
                        kind="text_fit",
                        severity=Severity.LOW,
                        message=(
                            f"文本纵向超出文本框：所需高 {text_h:.2f} 英寸 > "
                            f"可用 {avail_h:.2f} 英寸{v_approx}"
                        ),
                        primary=_shape_ref(rec),
                        details={
                            "text_height_in": round(text_h, 4),
                            "avail_height_in": round(avail_h, 4),
                            "word_wrap": bool(rec.word_wrap),
                        },
                        confidence=v_conf,
                    )
                )
        return findings


# ---------------------------------------------------------------- AutofitRisk


class AutofitRiskRule:
    rule_id = "autofit"

    def run(self, context: RuleContext) -> list[AuditFinding]:
        findings: list[AuditFinding] = []
        for rec in context.records:
            if rec.is_group or rec.is_connector or rec.is_hidden:
                continue
            if (
                rec.has_table
                or rec.geometry_unknown
                or not rec.has_text_frame
                or not rec.text.strip()
            ):
                continue
            if rec.role in _TEXT_FIT_SKIP_ROLES and _role_text_fit_ignored(
                rec.role, context.config
            ):
                continue
            if rec.autofit_norm_auto_fit:
                f = _shrink_finding(rec)
            elif rec.autofit_sp_auto_fit:
                f = _grow_finding(rec, context)
            else:
                continue
            if f is not None:
                findings.append(f)
        return findings


def _shrink_finding(rec: _ShapeRecord) -> AuditFinding | None:
    """normAutofit（缩小字体适应 Shape）：字号过小影响可读性。"""
    r = _rect(rec)
    avail_w = avail_h = 0.0
    if r is not None:
        avail_w = r.width - (rec.tf_margin_left or 0.0) - (rec.tf_margin_right or 0.0)
        avail_h = r.height - (rec.tf_margin_top or 0.0) - (rec.tf_margin_bottom or 0.0)
    over_w, over_h, _, _ = _text_overflow(rec, avail_w, avail_h)
    scale = rec.autofit_font_scale
    if scale is None and not (over_w or over_h):
        return None  # 未实际缩小也无溢出 → 无风险
    orig = _effective_font_size_pt(rec)
    est = (orig * scale) if (scale is not None and orig is not None) else None
    orig_txt = f"原始 {orig:g}pt" if orig is not None else "原始字号未记录"
    scale_txt = f"fontScale {scale:.0%}" if scale is not None else "fontScale 未记录"
    est_txt = f"，估算后 {est:.1f}pt" if est is not None else ""
    if est is not None and est < _MIN_READABLE_PT:
        severity = Severity.HIGH
        message = (
            f"字体缩小到 {est:.1f}pt，低于最小可读 {_MIN_READABLE_PT}pt（{orig_txt}，{scale_txt}）"
        )
    else:
        severity = Severity.MID
        message = f"文本被缩小字体适应 Shape（normAutofit）：{orig_txt}，{scale_txt}{est_txt}"
    return _finding(
        rule_id=RULE_AUTOFIT_SHRINK,
        kind="autofit",
        severity=severity,
        message=message,
        primary=_shape_ref(rec),
        details={
            "original_font_size_pt": orig,
            "font_scale": scale,
            "estimated_font_size_pt": est,
            "min_readable_pt": _MIN_READABLE_PT,
        },
        confidence=1.0 if scale is not None else 0.6,
    )


def _grow_finding(rec: _ShapeRecord, context: RuleContext) -> AuditFinding | None:
    """spAutoFit（扩大 Shape 适应文字）：撑大后越界/撞对象。"""
    r = _rect(rec)
    if r is None:
        return None
    avail_w = r.width - (rec.tf_margin_left or 0.0) - (rec.tf_margin_right or 0.0)
    avail_h = r.height - (rec.tf_margin_top or 0.0) - (rec.tf_margin_bottom or 0.0)
    over_w, over_h, text_w, text_h = _text_overflow(rec, avail_w, avail_h)
    if not (over_w or over_h):
        return None  # 现有框能装下 → 不会撑大
    page = context.page_rect
    tol = context.config.bounds_tolerance_in
    grown_right = r.right + max(text_w - avail_w, 0.0)
    grown_bottom = r.bottom + max(text_h - avail_h, 0.0)
    off_page = grown_right > page.right + tol or grown_bottom > page.bottom + tol
    parts = []
    if over_w:
        parts.append(f"宽 {text_w:.2f}>可用 {avail_w:.2f}")
    if over_h:
        parts.append(f"高 {text_h:.2f}>可用 {avail_h:.2f}")
    dims = "，".join(parts) + " 英寸"
    if off_page:
        severity = Severity.HIGH
        tail = "撑大后可能越出幻灯片"
    else:
        severity = Severity.MID
        tail = "撑大后可能撞到相邻对象"
    message = f"文本框按内容自动扩大（spAutoFit）：文本超出现有边界（{dims}），{tail}"
    return _finding(
        rule_id=RULE_AUTOFIT_GROW,
        kind="autofit",
        severity=severity,
        message=message,
        primary=_shape_ref(rec),
        details={
            "text_width_in": round(text_w, 4),
            "text_height_in": round(text_h, 4),
            "avail_width_in": round(avail_w, 4),
            "avail_height_in": round(avail_h, 4),
        },
        confidence=1.0,
    )


# ---------------------------------------------------------------- 编排


DEFAULT_RULES: list[AuditRule] = [
    BoundsRule(),
    MarginRule(),
    OverlapRule(),
    TextFitRule(),
    AutofitRiskRule(),
]


def _automatic_suppression_reason(
    record: _ShapeRecord, config: AuditConfig
) -> SuppressionReason | None:
    """按对应 ignore 开关决定角色豁免；未豁免返回 None。

    margin finding 的角色豁免唯一入口——角色不再无条件抑制，
    --no-…-ignore 关闭后同一 finding 恢复为普通 finding。
    """
    role = record.role
    if role == "background" and config.ignore_full_bleed_shapes:
        return "full_bleed"
    if role == "page_number" and config.ignore_page_numbers:
        return "page_number"
    if role in {"header", "footer"} and config.ignore_headers_footers:
        return "header_footer"
    if role == "decoration" and config.ignore_repeated_decorations:
        return "repeated_decoration"
    return None


def _role_text_fit_ignored(role: str, config: AuditConfig) -> bool:
    """文本拟合规则（TextFit/Autofit）是否跳过该角色（按对应 ignore 开关）。"""
    if role == "page_number":
        return config.ignore_page_numbers
    if role in {"header", "footer"}:
        return config.ignore_headers_footers
    return False


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
    """用户 ignore（primary/secondary 任一命中）优先；margin 角色豁免按开关门控。"""
    cfg = context.config
    refs = [f.primary]
    if f.secondary is not None:
        refs.append(f.secondary)
    for ref in refs:
        if cfg.ignored_shapes and (ref.slide_index, ref.shape_id) in cfg.ignored_shapes:
            return "user_shape"
        if cfg.ignored_regions:
            rec = _find_record(context.records, ref.slide_index, ref.shape_id)
            if rec is not None:
                r = _rect(rec)
                if r is not None:
                    cx, cy = r.center()
                    for rx, ry, rw, rh in cfg.ignored_regions:
                        if rx <= cx <= rx + rw and ry <= cy <= ry + rh:
                            return "user_region"
    rec = _find_record(context.records, f.primary.slide_index, f.primary.shape_id)
    if rec is not None and f.kind == "margin":
        return _automatic_suppression_reason(rec, cfg)
    return None
