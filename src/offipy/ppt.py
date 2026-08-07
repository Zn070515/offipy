"""PowerPoint 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActivePresentation 定位当前文稿。
"""

import math
import os
import shutil
import tempfile
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from typing import Any, NamedTuple

from . import core
from ._comguard import _COM_ERROR, guard_com, save_with_lock_retry
from .core import destructive, readonly_guard, requires_target
from .exceptions import (
    ComOperationError,
    FileConflictError,
    InvalidArgumentError,
    TargetNotFoundError,
)
from .models import (
    CoordinateSpace,
    ShapeInfo,
    SlideTextRecord,
    placeholder_type_name,
    shape_type_name,
)
from .paths import default_save_path, ensure_writable

PP_ALERTS_NONE = 1  # ppAlertsNone（=0 是 ppAlertsAll）
PP_FIXED_FORMAT_TYPE_PDF = 2  # ppFixedFormatTypePDF（ExportAsFixedFormat 的 OutputType）
PP_LAYOUT_TITLE = 1
PP_LAYOUT_TEXT = 2
PP_LAYOUT_TITLE_ONLY = 5
PP_LAYOUT_BLANK = 12

# PpPlaceholderType 官方值（微软 Learn，round-10 探针运行时常量 20/20 核实）
PP_PLACEHOLDER_TITLE = 1  # ppPlaceholderTitle
PP_PLACEHOLDER_BODY = 2  # ppPlaceholderBody
PP_PLACEHOLDER_CENTER_TITLE = 3  # ppPlaceholderCenterTitle
PP_PLACEHOLDER_SLIDE_NUMBER = 13  # ppPlaceholderSlideNumber
PP_PLACEHOLDER_HEADER = 14  # ppPlaceholderHeader
PP_PLACEHOLDER_FOOTER = 15  # ppPlaceholderFooter
PP_PLACEHOLDER_DATE = 16  # ppPlaceholderDate

# 无对应占位符时自动建文本框的默认位置（磅）：4:3 标准幻灯片
_TITLE_BOX = (36, 18, 648, 72)
_BODY_BOX = (36, 90, 648, 396)


def _placeholder_by_type(shapes, *pp_types):
    """按占位符类型找 shape（不硬编码 Placeholders(2) 序号）；找不到返回 None。"""
    placeholders = getattr(shapes, "Placeholders", None)
    if placeholders is None:
        return None
    for i in range(1, placeholders.Count + 1):
        if placeholders(i).PlaceholderFormat.Type in pp_types:
            return placeholders(i)
    return None


# ------------------------------------------------------------------ 读全（P1-4）


def _tri_state_to_bool(value) -> bool | None:
    """MsoTriState 正规化：-1/1→True、0→False、其余（-2/-3 混合态）→None。禁 bool() 偷译。"""
    if value in (-1, 1):
        return True
    if value == 0:
        return False
    return None


def _shape_has_text_frame(shape) -> bool:
    """shape 是否有文本能力：HasTextFrame 优先；None/读不到再兜底访问 TextFrame（P2-1）。"""
    try:
        state = _tri_state_to_bool(shape.HasTextFrame)
    except Exception:
        state = None
    if state is not None:
        return state
    try:
        shape.TextFrame  # noqa: B018 — 访问成功即证明有 TextFrame，忽略返回值
    except Exception:
        return False
    return True


MSO_GROUP = 6  # MsoShapeType.msoGroup
MSO_PLACEHOLDER = 14  # MsoShapeType.msoPlaceholder
MSO_FILL_SOLID = 1  # MsoFillType.msoFillSolid（仅 solid fill 给色）


def _shape_is_group(shape) -> bool:
    try:
        return int(shape.Type) == MSO_GROUP
    except Exception:
        return False


def _shape_is_rotated(shape) -> bool:
    """是否旋转（非 90° 整数倍）。旋转 group 内子元素读值不可信（探针 P0-2）。"""
    try:
        return float(shape.Rotation) % 90 != 0
    except Exception:
        return False


def _iter_shapes(
    shapes,
    *,
    recursive: bool,
    parent_shape_id: int | None = None,
    group_path: tuple[int, ...] = (),
    rotated: bool = False,
):
    """统一遍历器：产出 (shape, parent_shape_id, group_path, rotated)（v0.12 read_shapes 共用）。

    - top-level：parent_shape_id=None、group_path=()；group 子元素 parent_shape_id=直接父
      group 的 shape_id、group_path=祖先链（外层→内层）。
    - rotated：shape 是否处于**旋转 group 内**（非旋转 group 子元素读值是幻灯片绝对坐标；
      旋转 group 子元素读值不可信 → coordinate_space="unknown"）。
    - 先判 Type==6 再访问 GroupItems：COM Group() 会把嵌套 group 拍平，对拍平成员直接
      访问 GroupItems 抛 E_ACCESSDENIED（探针实证）。
    """
    try:
        count = int(shapes.Count)
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            shape = shapes(i)
        except Exception:
            continue
        sid = _shape_id(shape)
        is_group = _shape_is_group(shape)
        yield (shape, parent_shape_id, group_path, rotated)
        if is_group and recursive:
            try:
                items = shape.GroupItems
            except Exception:
                continue
            child_rotated = rotated or _shape_is_rotated(shape)
            yield from _iter_shapes(
                items,
                recursive=True,
                parent_shape_id=sid,
                group_path=group_path + (sid,),
                rotated=child_rotated,
            )


def _shape_id(shape) -> int:
    try:
        return int(shape.Id)
    except Exception:
        return 0


def _shape_name(shape) -> str:
    try:
        return str(shape.Name)
    except Exception:
        return ""


def _shape_text(shape) -> str:
    try:
        return str(shape.TextFrame.TextRange.Text)
    except Exception:
        return ""


def _shape_float(shape, attr: str) -> float:
    try:
        return float(getattr(shape, attr))
    except Exception:
        return 0.0


def _shape_z_order(shape) -> int:
    """ZOrderPosition 兜底大数：读不到排最后（稳定，不干扰阅读顺序）。"""
    try:
        return int(shape.ZOrderPosition)
    except Exception:
        return 1_000_000


def _placeholder_info(shape) -> tuple[bool, int | None, str | None]:
    """(is_placeholder, type, type_name)。shape.Type==14（msoPlaceholder）判定占位符。"""
    try:
        if int(shape.Type) != MSO_PLACEHOLDER:
            return False, None, None
        ph_type = shape.PlaceholderFormat.Type
        if ph_type is None:
            return True, None, None
        ph_type = int(ph_type)
        return True, ph_type, placeholder_type_name(ph_type)
    except Exception:
        return False, None, None


def _require_shape_id(shape) -> int:
    """严格 shape 身份：读不到 Id 抛 ComOperationError（read_shapes 无 0 兜底）。"""
    try:
        return int(shape.Id)
    except Exception as e:
        raise ComOperationError(f"shape.Id 读取失败（无稳定身份）: {e}") from e


def _rgb_to_hex(rgb: int) -> str:
    """COM RGB（BGR 打包）→ #RRGGBB。"""
    rgb = int(rgb) & 0xFFFFFF
    r = rgb & 0xFF
    g = (rgb >> 8) & 0xFF
    b = (rgb >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def _shape_type(shape) -> tuple[int, str]:
    """(MsoShapeType 数值, 名称)；读不到 Type → (0, "unknown_0")。"""
    try:
        t = int(shape.Type)
    except Exception:
        t = 0
    return t, shape_type_name(t)


def _shape_rotation(shape) -> float:
    try:
        return float(shape.Rotation)
    except Exception:
        return 0.0


def _shape_visible(shape) -> bool | None:
    """MsoTriState 正规化：-1/1→True、0→False、混合态→None。"""
    try:
        return _tri_state_to_bool(shape.Visible)
    except Exception:
        return None


def _shape_fill(shape) -> tuple[str | None, float | None]:
    """(hex_color, transparency)。仅 solid fill 给色；gradient/pattern/picture → None。"""
    try:
        fill = shape.Fill
        if int(fill.Type) != MSO_FILL_SOLID:
            return None, None
    except Exception:
        return None, None
    color = None
    try:
        color = _rgb_to_hex(fill.ForeColor.RGB)
    except Exception:
        color = None
    transparency = None
    try:
        transparency = float(fill.Transparency)
    except Exception:
        transparency = None
    return color, transparency


def _shape_line(shape) -> tuple[str | None, float | None]:
    """(hex_color, width)。仅可见 line 给色；隐藏 line → 均 None。"""
    try:
        line = shape.Line
        visible = _tri_state_to_bool(line.Visible)
    except Exception:
        return None, None
    if visible is False:
        return None, None
    color = None
    try:
        color = _rgb_to_hex(line.ForeColor.RGB)
    except Exception:
        color = None
    width = None
    try:
        width = float(line.Weight)
    except Exception:
        width = None
    return color, width


def _shape_font(shape) -> tuple[float | None, str | None, str | None]:
    """(size, name, color_hex) 首 run 语义；无文本/空文本 → 全 None。"""
    if not _shape_has_text_frame(shape):
        return None, None, None
    try:
        tr = shape.TextFrame.TextRange
    except Exception:
        return None, None, None
    try:
        if not str(tr.Text):
            return None, None, None
    except Exception:
        return None, None, None
    try:
        font = tr.Runs(1).Font
    except Exception:
        return None, None, None
    size = name = color = None
    try:
        size = float(font.Size)
    except Exception:
        size = None
    try:
        name = str(font.Name)
    except Exception:
        name = None
    try:
        color = _rgb_to_hex(font.Color.RGB)
    except Exception:
        color = None
    return size, name, color


def _local_z_order_rank(shapes) -> list[int]:
    """所在集合内每 shape 的 1-based z-order rank（与 shapes 元素顺序对齐）。

    顶层（slide.Shapes）ZOrderPosition 即 1..Count，rank 恒等于 ZOrderPosition；
    group 子元素 ZOrderPosition 带偏移（探针 #6），按兄弟 ZOrderPosition 排序推出的
    rank 还原为 group-local 1..Count。读不到 ZOrderPosition 的排最后（稳定兜底）。
    """
    zs = [_shape_z_order(sh) for sh in shapes]
    return [1 + sum(1 for oz in zs if oz < z) for z in zs]


def _iter_shape_records(
    shapes,
    *,
    recursive: bool,
    parent_shape_id: int | None = None,
    group_path: tuple[int, ...] = (),
    rotated: bool = False,
):
    """read_shapes 专用遍历：产出 (shape, parent_shape_id, group_path, rotated, z_order)。

    - shape_id 严格：任何 shape Id 读不到 → _require_shape_id 抛 ComOperationError。
    - z_order 为所在集合内 1-based rank（_local_z_order_rank）。
    - rotated：是否处于旋转 group 内 → coordinate_space="unknown"。
    - 先判 Type==6 再访问 GroupItems（COM Group() 拍平嵌套 group 时直接访问抛错，探针实证）。
    """
    items: list = []
    try:
        count = int(shapes.Count)
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            items.append(shapes(i))
        except Exception:
            continue
    ranks = _local_z_order_rank(items)
    for shape, z_order in zip(items, ranks, strict=True):
        sid = _require_shape_id(shape)
        is_group = _shape_is_group(shape)
        yield (shape, parent_shape_id, group_path, rotated, z_order)
        if is_group and recursive:
            try:
                group_items = shape.GroupItems
            except Exception:
                continue
            child_rotated = rotated or _shape_is_rotated(shape)
            yield from _iter_shape_records(
                group_items,
                recursive=True,
                parent_shape_id=sid,
                group_path=group_path + (sid,),
                rotated=child_rotated,
            )


def _record_shape_info(
    shape,
    *,
    parent_shape_id: int | None,
    group_path: tuple[int, ...],
    rotated: bool,
    z_order: int,
) -> ShapeInfo:
    """读 shape → ShapeInfo。可选字段逐 try/except 兜底（诚实读，不编假值）。"""
    is_ph, ph_type, ph_name = _placeholder_info(shape)
    has_text = _shape_has_text_frame(shape)
    font_size, font_name, font_color = _shape_font(shape)
    fill_color, fill_transparency = _shape_fill(shape)
    line_color, line_width = _shape_line(shape)
    st, st_name = _shape_type(shape)
    coordinate_space: CoordinateSpace = "unknown" if rotated else "slide"
    return {
        "shape_id": _require_shape_id(shape),
        "name": _shape_name(shape),
        "shape_type": st,
        "shape_type_name": st_name,
        "left": _shape_float(shape, "Left"),
        "top": _shape_float(shape, "Top"),
        "width": _shape_float(shape, "Width"),
        "height": _shape_float(shape, "Height"),
        "coordinate_space": coordinate_space,
        "coordinate_unit": "pt",
        "rotation": _shape_rotation(shape),
        "visible": _shape_visible(shape),
        "fill_color": fill_color,
        "fill_transparency": fill_transparency,
        "line_color": line_color,
        "line_width": line_width,
        "has_text_frame": has_text,
        "text": _shape_text(shape),
        "font_size": font_size,
        "font_name": font_name,
        "font_color": font_color,
        "is_placeholder": is_ph,
        "placeholder_type": ph_type,
        "placeholder_type_name": ph_name,
        "parent_shape_id": parent_shape_id,
        "group_path": list(group_path),
        "z_order": z_order,
    }


def _record_from_shape(
    shape,
    *,
    parent_shape_id: int | None,
    group_path: tuple[int, ...],
    rotated: bool,
) -> SlideTextRecord:
    """读 shape → SlideTextRecord；逐属性 try/except 兜底。坐标单位恒为磅（pt）。"""
    is_ph, ph_type, ph_name = _placeholder_info(shape)
    coordinate_space: CoordinateSpace = "unknown" if rotated else "slide"
    return {
        "shape_id": _shape_id(shape),
        "name": _shape_name(shape),
        "text": _shape_text(shape),
        "left": _shape_float(shape, "Left"),
        "top": _shape_float(shape, "Top"),
        "width": _shape_float(shape, "Width"),
        "height": _shape_float(shape, "Height"),
        "coordinate_space": coordinate_space,
        "coordinate_unit": "pt",
        "is_placeholder": is_ph,
        "placeholder_type": ph_type,
        "placeholder_type_name": ph_name,
        "parent_shape_id": parent_shape_id,
        "group_path": list(group_path),
    }


class _InternalTextShapeRecord(NamedTuple):
    """公开 SlideTextRecord + 内部 z_order（P1-4：阅读排序用，不进公开模型）。"""

    record: SlideTextRecord
    z_order: int


def _collect_text_records(slide, *, recursive: bool = True) -> list[_InternalTextShapeRecord]:
    """收集 slide 上全部**有文本能力**的 shape：[(record, z_order)]。"""
    out: list[_InternalTextShapeRecord] = []
    for shape, parent_id, group_path, rotated in _iter_shapes(slide.Shapes, recursive=recursive):
        if not _shape_has_text_frame(shape):
            continue
        out.append(
            _InternalTextShapeRecord(
                _record_from_shape(
                    shape,
                    parent_shape_id=parent_id,
                    group_path=group_path,
                    rotated=rotated,
                ),
                _shape_z_order(shape),
            )
        )
    return out


def _reading_order_key(item: _InternalTextShapeRecord):
    """稳定阅读顺序（P1-4）：top 按 5pt 一档取整（floor 防银行家舍入）→ left → z → id。"""
    return (
        math.floor((item.record["top"] + 2.5) / 5.0),
        item.record["left"],
        item.z_order,
        item.record["shape_id"],
    )


def _page_size_pt(pres) -> tuple[float, float]:
    """演示文稿页面尺寸（磅）(width, height)；读不到回宽屏 16:9 默认 (960, 540)。"""
    try:
        ps = pres.PageSetup
        return float(ps.SlideWidth), float(ps.SlideHeight)
    except Exception:
        return 960.0, 540.0


def _require_slide(pres, slide_idx: int):
    """slide 索引前置校验：1..演示文稿页数，越界抛 InvalidArgumentError。

    把 PowerPoint 原生枚举错误（"Slides.Item: Integer out of range"）前置成
    offipy 语义化异常，脚本里 slide 号算错一位时直接得到可读提示。
    """
    count = pres.Slides.Count
    if slide_idx < 1 or slide_idx > count:
        raise InvalidArgumentError(f"slide {slide_idx} 越界，演示文稿共 {count} 页")
    return pres.Slides(slide_idx)


# 摘要豁免占位符类型（P1-2）：页码/页眉/页脚/日期，不进 title/body
_EXEMPT_PLACEHOLDER_TYPES = frozenset({13, 14, 15, 16})
_PAGE_NUMBER_MAX_WIDTH_PT = 72.0  # 页码通常远小于 72pt 宽


def _is_page_number_candidate(rec: SlideTextRecord, pw: float, ph: float) -> bool:
    """页码候选：纯数字文本 AND 位于页面底部/右下角落 AND 宽度较小（P1-2）。

    只豁免「极像页码」的文本；普通纯数字（年份/章节号/KPI）不豁免。
    """
    text = rec["text"].strip()
    if not text.isdigit():
        return False
    bottom = rec["top"] > 0.8 * ph
    corner = rec["left"] > 0.9 * pw and rec["top"] > 0.7 * ph
    if not (bottom or corner):
        return False
    return rec["width"] <= _PAGE_NUMBER_MAX_WIDTH_PT


def _is_exempt_text(rec: SlideTextRecord, pw: float, ph: float) -> bool:
    """摘要豁免集：页码/页眉/页脚/日期占位符 + 页码候选。"""
    if rec["is_placeholder"] and rec["placeholder_type"] in _EXEMPT_PLACEHOLDER_TYPES:
        return True
    return _is_page_number_candidate(rec, pw, ph)


def _read_notes(slide) -> str:
    """读取演讲者备注文本；无正文占位符/读取失败回空串。"""
    try:
        ph = _placeholder_by_type(slide.NotesPage.Shapes, PP_PLACEHOLDER_BODY)
        if ph is None:
            return ""
        return str(ph.TextFrame.TextRange.Text)
    except Exception:
        return ""


def _summarize_slide(slide, index: int, pw: float, ph: float) -> dict:
    """单页摘要：title/body 启发式聚合 + notes（兼容 0.9 read_slide_texts 语义）。"""
    items = _collect_text_records(slide, recursive=True)
    title_ph = body_ph = None
    for item in items:
        rec = item.record
        if not rec["is_placeholder"]:
            continue
        t = rec["placeholder_type"]
        if title_ph is None and t in (PP_PLACEHOLDER_TITLE, PP_PLACEHOLDER_CENTER_TITLE):
            title_ph = rec
        if body_ph is None and t == PP_PLACEHOLDER_BODY:
            body_ph = rec
    used = {
        sid
        for sid in (
            title_ph["shape_id"] if title_ph else None,
            body_ph["shape_id"] if body_ph else None,
        )
        if sid is not None
    }
    if title_ph is not None:
        title = title_ph["text"]
    else:
        cands = [
            item
            for item in items
            if item.record["shape_id"] not in used
            and not _is_exempt_text(item.record, pw, ph)
            and item.record["text"]
        ]
        if cands:
            first = min(cands, key=_reading_order_key)
            title = first.record["text"]
            used.add(first.record["shape_id"])
        else:
            title = ""
    if body_ph is not None:
        body = body_ph["text"]
    else:
        body = "\n".join(
            item.record["text"]
            for item in sorted(items, key=_reading_order_key)
            if item.record["shape_id"] not in used
            and not _is_exempt_text(item.record, pw, ph)
            and item.record["text"]
        )
    return {"index": index, "title": title, "body": body, "notes": _read_notes(slide)}


# ------------------------------------------------------------------ 编辑定位 / 校验（S1）


@dataclass
class _LocatedShape:
    """编辑操作定位结果：shape 本身 + 所在集合（slide.Shapes 或父 GroupItems）。

    - containing_collection：z-order / 删除在正确的集合内操作（group 子元素在父
      GroupItems，绝不跨集合）。
    - parent_shape_id / group_path：与 read_shapes 的语义一致（外层→内层）。
    - rotated_group_ancestor：是否处于旋转 group 内（几何读值不可信）。
    """

    shape: Any
    containing_collection: Any
    parent_shape_id: int | None
    group_path: tuple[int, ...]
    rotated_group_ancestor: bool


def _locate_in_shapes(
    shapes,
    shape_id: int,
    *,
    parent_shape_id: int | None,
    group_path: tuple[int, ...],
    rotated: bool,
) -> _LocatedShape | None:
    """在 shapes 集合内按 shape_id 递归找 shape；未命中返回 None。

    shapes 是 slide.Shapes 或某 group 的 GroupItems；parent_shape_id/group_path 描述
    shapes 的**子元素**的父级关系（top-level → (None, ())；group 子元素 → 父 id +
    祖先链）。shape_id 严格：遍历途中任何 Id 读不到 → ComOperationError（与
    read_shapes 一致）。绝不调用 Shapes.Range([id])（COM Range 按名称/序号索引，
    id 匹配不可靠，探针 #4 建议逐 shape 扫描）。
    """
    try:
        count = int(shapes.Count)
    except Exception:
        return None
    for i in range(1, count + 1):
        try:
            shape = shapes(i)
        except Exception:
            continue
        sid = _require_shape_id(shape)
        if sid == shape_id:
            return _LocatedShape(
                shape=shape,
                containing_collection=shapes,
                parent_shape_id=parent_shape_id,
                group_path=group_path,
                rotated_group_ancestor=rotated,
            )
        if _shape_is_group(shape):
            try:
                items = shape.GroupItems
            except Exception:
                continue
            child_rotated = rotated or _shape_is_rotated(shape)
            hit = _locate_in_shapes(
                items,
                shape_id,
                parent_shape_id=sid,
                group_path=group_path + (sid,),
                rotated=child_rotated,
            )
            if hit is not None:
                return hit
    return None


def _find_shape_by_id(slide, shape_id: int) -> _LocatedShape:
    """递归定位 shape（顶层 + group 后代）；找不到 → TargetNotFoundError。"""
    hit = _locate_in_shapes(
        slide.Shapes, shape_id, parent_shape_id=None, group_path=(), rotated=False
    )
    if hit is None:
        raise TargetNotFoundError(
            f"shape {shape_id} 不存在于该页（shape_id 是 PowerPoint 的 Shape.Id；"
            f"用 read_shapes 核对当前页 shape_id 列表）"
        )
    return hit


def _validate_hex_color(value, name: str = "color") -> str:
    """严格 #RRGGBB（6 位十六进制）；非法抛 InvalidArgumentError。返回规范大写。"""
    if not isinstance(value, str):
        raise InvalidArgumentError(f"{name} 必须是 #RRGGBB 字符串，收到 {type(value).__name__}")
    s = value.strip().upper()
    if len(s) != 7 or not s.startswith("#"):
        raise InvalidArgumentError(f"{name} 必须是 #RRGGBB 格式（7 字符含 #），收到 {value!r}")
    try:
        int(s[1:], 16)
    except ValueError:
        raise InvalidArgumentError(f"{name} 含非法十六进制字符: {value!r}") from None
    return s


def _rgb_to_com(hex_color: str) -> int:
    """#RRGGBB → COM RGB（BGR 打包，低字节 R）。"""
    s = hex_color.lstrip("#")
    r = int(s[0:2], 16)
    g = int(s[2:4], 16)
    b = int(s[4:6], 16)
    return r | (g << 8) | (b << 16)


def _validate_fraction_0_1(value, name: str = "transparency") -> float:
    """透明度 [0,1]；非法抛 InvalidArgumentError。"""
    try:
        f = float(value)
    except (TypeError, ValueError):
        raise InvalidArgumentError(f"{name} 必须是数值，收到 {value!r}") from None
    if not math.isfinite(f) or f < 0.0 or f > 1.0:
        raise InvalidArgumentError(f"{name} 必须在 [0,1] 内，收到 {value!r}")
    return f


def _validate_positive_float(value, name: str) -> float:
    """> 0 的有限正数（width/height/font size）；非法抛 InvalidArgumentError。"""
    try:
        f = float(value)
    except (TypeError, ValueError):
        raise InvalidArgumentError(f"{name} 必须是数值，收到 {value!r}") from None
    if not math.isfinite(f) or f <= 0.0:
        raise InvalidArgumentError(f"{name} 必须是 > 0 的有限正数，收到 {value!r}")
    return f


def _validate_finite_float(value, name: str) -> float:
    """有限数值（坐标/旋转）；NaN/Inf 抛 InvalidArgumentError。"""
    try:
        f = float(value)
    except (TypeError, ValueError):
        raise InvalidArgumentError(f"{name} 必须是数值，收到 {value!r}") from None
    if not math.isfinite(f):
        raise InvalidArgumentError(f"{name} 必须是有限数值，收到 {value!r}")
    return f


def _require_text_frame(shape, op_desc: str) -> None:
    """shape 必须有文本能力；图片/线条/无文本图形 → InvalidArgumentError。"""
    if not _shape_has_text_frame(shape):
        raise InvalidArgumentError(f"{op_desc} 需要文本能力，但目标 shape 无 TextFrame")


def _require_fill_capability(shape, op_desc: str) -> None:
    """shape 必须支持 Fill；读不到 Fill 对象 → InvalidArgumentError。"""
    try:
        shape.Fill  # noqa: B018 — 访问成功即证明有 Fill 能力，忽略返回值
    except Exception:
        raise InvalidArgumentError(f"{op_desc} 需要填充能力，但目标 shape 不支持 Fill") from None


def _require_line_capability(shape, op_desc: str) -> None:
    """shape 必须支持 Line；读不到 Line 对象 → InvalidArgumentError。"""
    try:
        shape.Line  # noqa: B018 — 访问成功即证明有 Line 能力，忽略返回值
    except Exception:
        raise InvalidArgumentError(f"{op_desc} 需要轮廓能力，但目标 shape 不支持 Line") from None


@guard_com
class PptApp:
    def __init__(self, visible: bool = True, modify_existing_visibility: bool = False):
        self.app, self.created = core.ensure_app(
            "ppt", visible=visible, modify_existing_visibility=modify_existing_visibility
        )
        # _owned：本库启动的实例才允许 quit() 直接退出；连到既有实例默认拒绝
        self._owned = self.created
        # DisplayAlerts 不再永久静音（P0-5）：按需用 _alerts_scope 临时抑制
        self._saved_alerts = self.app.DisplayAlerts  # quit() 兜底还原
        # 记录本实例进程 PID（断连自愈/退出时精确清理，不误杀用户其它实例）
        self._pid = core.app_process_pid(self.app, "ppt")
        self._docs: dict[str, Any] = {}  # doc_id → 演示文稿句柄（P2-2 多文档）
        self._active_id: str | None = None
        self._seq = 0

    @contextmanager
    def _alerts_scope(self, value: int = PP_ALERTS_NONE):
        """临时抑制模态对话框；退出时（含异常路径）还原 DisplayAlerts 原值。"""
        prev = self.app.DisplayAlerts
        self.app.DisplayAlerts = value
        try:
            yield
        finally:
            self.app.DisplayAlerts = prev

    def _stable_identity(self, obj):
        """稳定身份键（P0-4）：已保存 → (FullName.lower(), None)；未保存 → (None, Name.lower())。"""
        try:
            fullname = obj.FullName
        except Exception:
            fullname = None
        try:
            name = obj.Name
        except Exception:
            name = None
        try:
            path = obj.Path
        except Exception:
            path = None
        if path:
            return (str(fullname).lower() if fullname else None, None)
        return (None, name.lower() if name else None)

    def _register(self, obj) -> str:
        """登记新文档句柄，分配 doc_id 并设为活动；同底层文档复用已有 doc_id。"""
        ident = self._stable_identity(obj)
        if ident != (None, None):
            for did, pres in self._docs.items():
                if self._stable_identity(pres) == ident:
                    self._docs[did] = obj  # 复用 doc_id，换用实时句柄
                    self._active_id = did
                    return did
        self._seq += 1
        did = f"pres{self._seq}"
        self._docs[did] = obj
        self._active_id = did
        return did

    def _sync_registered(self, obj) -> str:
        """把实时解析到的句柄并入文档表：已登记则复用并置活动，否则登记为新文档。"""
        for did, pres in self._docs.items():
            if pres is obj:
                self._active_id = did
                return did
        return self._register(obj)

    # --- 演示文稿（P2-2 多文档：doc_id 显式路由，缺省走活动） ---
    def new_pres(self) -> str:
        """新建空白演示文稿，登记进文档表并设为活动。返回 doc_id。"""
        return self._register(self.app.Presentations.Add())

    def open_pres(self, path: str) -> str:
        """打开现有演示文稿并设为活动。返回 doc_id。"""
        return self._register(self.app.Presentations.Open(os.path.abspath(path)))

    @destructive
    def close_pres(self, save: bool = True, doc_id: str | None = None):
        """关闭演示文稿（doc_id 缺省为活动），不退出 PowerPoint。

        save=True → 先保存（从未保存过则自动落盘用户数据目录，不弹另存为）并返回
        保存路径；save=False → 直接关闭不保存、不弹对话框，返回 None。
        语义对齐 word close_doc / excel close_book（#26：Ppt 此前只有 quit）。
        """
        pres = self._require_pres(doc_id)
        did = doc_id if doc_id is not None else self._active_id
        if save:
            path = pres.FullName if pres.Path else self.save(doc_id=did)
            with self._alerts_scope():
                pres.Close()
        else:
            path = None
            with self._alerts_scope():
                pres.Saved = True  # 兜底：确保 Close 不触发保存提示
                pres.Close()
        if did is not None:
            self._docs.pop(did, None)
            if self._active_id == did:
                self._active_id = None
        return path

    def active_pres(self, doc_id: str | None = None):
        # 显式 doc_id：绑定目标路由，只查文档表；未知/失效句柄抛 TargetNotFoundError。
        # 缺省 active：实时解析 ActivePresentation（doc_id 权威——绝不静默用陈旧的
        # _active_id 快路径，防「用户看到 B、Agent 以为 A」），解析到即并入文档表。
        # P0-8：全程纯探测，绝不隐式 Presentations.Add()。
        if doc_id is not None:
            pres = self._docs.get(doc_id)
            if pres is None or not core.doc_alive(pres):
                raise TargetNotFoundError(
                    f"未知演示文稿句柄: {doc_id!r}（当前会话未打开；doc_id 只在同会话内有效，"
                    f"本地直连 Ppt() 与会话式 Remote*/CLI/HTTP 互不相通；用本会话 list_docs 核对）"
                )
            return pres
        pres = core.active_doc("ppt", "ActivePresentation")
        if pres is not None:
            self._sync_registered(pres)
            return pres
        pres = self.app.ActivePresentation
        if pres is None:
            return None
        self._sync_registered(pres)
        return pres

    def _require_pres(self, doc_id: str | None = None):
        """操作前置：目标演示文稿不存在则抛 TargetNotFoundError，不隐式创建。"""
        pres = self.active_pres(doc_id)
        if pres is None:
            raise TargetNotFoundError("没有打开的演示文稿，请先 new_pres/open_pres")
        return pres

    def activate(self, doc_id: str) -> str:
        """把指定文档设为活动目标并同步真实 UI；未知句柄抛 TargetNotFoundError。"""
        pres = self._docs.get(doc_id)
        if pres is None or not core.doc_alive(pres):
            raise TargetNotFoundError(
                f"未知演示文稿句柄: {doc_id!r}（当前会话未打开；doc_id 只在同会话内有效，"
                f"本地直连 Ppt() 与会话式 Remote*/CLI/HTTP 互不相通；用本会话 list_docs 核对）"
            )
        old = self._active_id
        self._active_id = doc_id
        try:
            # PowerPoint 的 Presentation 无 Activate：激活其文档窗口；失败兜底
            try:
                pres.Windows.Item(1).Activate()
            except Exception:
                pres.Activate()
        except Exception as e:
            self._active_id = old  # 同步不上则回滚，不静默假活
            raise ComOperationError(f"激活演示文稿 {doc_id} 失败: {e}") from e
        return doc_id

    def list_docs(self) -> dict:
        """当前打开的文档表：{doc_id: {"name", "path", "active"}}。只报已登记句柄，不隐式枚举。

        P1-5：先并入真实活动焦点（active_pres 解析 ActivePresentation 进文档表、
        刷新 _active_id），active 标记跟随用户当前看到的文稿，不报陈旧焦点。
        """
        with suppress(Exception):
            self.active_pres()
        out = {}
        for did, pres in self._docs.items():
            if not core.doc_alive(pres):
                continue
            try:
                name = pres.Name
            except Exception:
                name = None
            try:
                path = pres.FullName
            except Exception:
                path = None
            out[did] = {"name": name, "path": path, "active": did == self._active_id}
        return out

    def get_target(self, doc_id: str | None = None):
        """目标身份 {app, doc_id, name, path}；无目标返回 None。只读探测。

        显式 doc_id：只查文档表，未注册/失效抛 TargetNotFoundError；
        缺省：当前活动目标。
        """
        if doc_id is not None:
            pres = self.active_pres(doc_id)
            resolved = doc_id
        else:
            pres = self.active_pres()
            if pres is None:
                return None
            active_id = self._active_id
            assert active_id is not None  # active_pres 非 None 时活动 id 必已同步
            resolved = active_id
        try:
            name = pres.Name
        except Exception:
            name = None
        try:
            path = pres.FullName
        except Exception:
            path = None
        return {"app": "ppt", "doc_id": resolved, "name": name, "path": path}

    @destructive
    def save(self, path: str | None = None, overwrite: bool = False, doc_id: str | None = None):
        """保存演示文稿并返回绝对路径。

        给 path → 另存到该路径；未给 path → 已保存过的存回原路径，从未保存过的
        自动落盘 <用户数据目录>/documents/<名字>_<时间戳>.pptx（不弹另存为对话框）。
        """
        if path:
            dest = ensure_writable(path, overwrite)  # 覆盖保护先于触 COM（fail-fast）
            pres = self._require_pres(doc_id)
            with self._alerts_scope():
                save_with_lock_retry(lambda: pres.SaveAs(dest), what="保存演示文稿")
            return dest
        pres = self._require_pres(doc_id)
        with self._alerts_scope():
            if pres.Path:  # 已有保存路径 → 原位保存
                pres.Save()
                return pres.FullName
            dest = default_save_path(pres.Name, ".pptx")
            pres.SaveAs(dest)
            return dest

    @requires_target
    def save_pdf(self, path: str, overwrite: bool = False, doc_id: str | None = None):
        dest = ensure_writable(path, overwrite)
        # ExportAsFixedFormat 第 2 参数是必填的 FixedFormatType（PDF=2）；Intent
        # 是打印品质（打印=2）；OutputType 默认 Slides=1（导出全部幻灯片）。
        # PrintRange 是 VT_DISPATCH 槽位，必须显式 None——makepy 生成的默认值
        # 0 是 int，直接塞进 dispatch 槽会 COM 转换失败。
        with self._alerts_scope():
            save_with_lock_retry(
                lambda: self._require_pres(doc_id).ExportAsFixedFormat(
                    dest, FixedFormatType=PP_FIXED_FORMAT_TYPE_PDF, Intent=2, PrintRange=None
                ),
                what="导出 PDF",
            )

    @requires_target
    def export_slides(
        self,
        out_dir: str,
        width: int = 1920,
        height: int = 1080,
        overwrite: bool = False,
        doc_id: str | None = None,
    ):
        """把当前演示文稿每一页导出为 PNG，供 Claude 视觉迭代。

        默认拒绝覆盖已有输出；overwrite=True 时先导出到同卷 staging 临时目录，
        全部成功后 os.replace 原子替换，中途失败不留半成品。
        """
        out_dir = os.path.abspath(out_dir)
        pres = self._require_pres(doc_id)
        count = pres.Slides.Count
        targets = [os.path.join(out_dir, f"slide_{i:02d}.png") for i in range(1, count + 1)]
        if not overwrite:
            existing = [p for p in targets if os.path.exists(p)]
            if existing:
                raise FileConflictError(f"导出目标已存在: {existing[0]}（overwrite=True 覆盖）")
        os.makedirs(out_dir, exist_ok=True)
        staging = tempfile.mkdtemp(prefix=".offipy-slides-", dir=os.path.dirname(out_dir) or ".")
        try:
            tmp_paths = []
            for i in range(1, count + 1):
                tmp = os.path.join(staging, f"slide_{i:02d}.png")
                pres.Slides(i).Export(tmp, "PNG", width, height)
                tmp_paths.append(tmp)
            for tmp, final in zip(tmp_paths, targets, strict=True):
                os.replace(tmp, final)
        finally:
            shutil.rmtree(staging, ignore_errors=True)
        return targets

    # --- 幻灯片 ---
    @destructive
    def add_slide(self, layout: int = PP_LAYOUT_TEXT, doc_id: str | None = None):
        if isinstance(layout, bool) or not isinstance(layout, int):
            raise InvalidArgumentError(f"非法 layout: {layout!r}（期望整数，如 1/2/5/12）")
        if layout < 1:
            raise InvalidArgumentError(f"非法 layout: {layout}（期望 ≥ 1）")
        pres = self._require_pres(doc_id)
        try:
            pres.Slides.Add(pres.Slides.Count + 1, layout)
        except _COM_ERROR:
            raise InvalidArgumentError(
                f"非法 layout: {layout}（当前模板不提供该布局；本机实测 1/2/5/12 可用）"
            ) from None
        return pres.Slides.Count

    @destructive
    def set_title(self, slide_idx: int, text: str, doc_id: str | None = None):
        if not text:
            raise InvalidArgumentError("set_title: text 不能为空")
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        ph = _placeholder_by_type(slide.Shapes, PP_PLACEHOLDER_TITLE, PP_PLACEHOLDER_CENTER_TITLE)
        if ph is None:
            ph = slide.Shapes.AddTextbox(1, *_TITLE_BOX)
        ph.TextFrame.TextRange.Text = text
        return ph.Id

    @destructive
    def set_body(self, slide_idx: int, lines, doc_id: str | None = None):
        if isinstance(lines, str):
            lines = [lines]
        if not lines:
            raise InvalidArgumentError("set_body: lines 不能为空")
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        ph = _placeholder_by_type(slide.Shapes, PP_PLACEHOLDER_BODY)
        if ph is None:
            ph = slide.Shapes.AddTextbox(1, *_BODY_BOX)
        ph.TextFrame.TextRange.Text = "\r".join(lines)
        return ph.Id

    @destructive
    def set_notes(self, slide_idx: int, text: str, doc_id: str | None = None):
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        shapes = slide.NotesPage.Shapes
        ph = _placeholder_by_type(shapes, PP_PLACEHOLDER_BODY)
        if ph is None:
            ph = shapes.AddTextbox(1, *_BODY_BOX)
        ph.TextFrame.TextRange.Text = text
        return ph.Id

    @destructive
    def add_textbox(
        self,
        slide_idx: int,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str,
        doc_id: str | None = None,
    ):
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        tb = slide.Shapes.AddTextbox(1, left, top, width, height)
        tb.TextFrame.TextRange.Text = text

    @destructive
    def add_picture(
        self,
        slide_idx: int,
        path: str,
        left: float,
        top: float,
        width: float,
        height: float,
        doc_id: str | None = None,
    ):
        if not os.path.isfile(path):
            raise InvalidArgumentError(f"源文件不存在: {path}")
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        # SaveWithDocument=msoTrue(-1)：LinkToFile=False 时必须内嵌，传 0 会被
        # PowerPoint 拒为 E_INVALIDARG。路径 normpath 归一化，COM 拒收正斜杠。
        slide.Shapes.AddPicture(
            os.path.normpath(os.path.abspath(path)), 0, -1, left, top, width, height
        )

    @readonly_guard
    def read_slide_texts(
        self,
        slide_idx: int,
        *,
        include_empty: bool = False,
        recursive: bool = True,
        doc_id: str | None = None,
    ) -> list[SlideTextRecord]:
        """读取第 slide_idx 页全部**具有文本能力**的 shape 文本（含 group 内文本）。

        - 只返回有 TextFrame 的 shape；图片/线条/无文本图形不在此列（read_shapes 的职责）。
        - include_empty=True 连文本为空的 TextFrame shape 也返回；False 只返回文本非空。
        - recursive=True 递归 group；坐标单位恒为磅（pt），coordinate_space 按探针结论
          标注（非旋转 group 子元素为幻灯片绝对坐标 "slide"；旋转 group 内不可信 "unknown"）。
        """
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        return [
            item.record
            for item in _collect_text_records(slide, recursive=recursive)
            if include_empty or item.record["text"]
        ]

    @readonly_guard
    def read_slide_summary(self, doc_id: str | None = None) -> list[dict]:
        """逐页读标题/正文/备注摘要（0.9 read_slide_texts 的语义），返回 list[dict]。

        - title：标题/居中标题占位符（type 1/3）优先；否则按稳定阅读顺序回退第一个非豁免文本。
        - body：正文占位符（type 2）优先；否则其余文本 shape 按阅读顺序 "\\n" 拼接。
        - 豁免集：页码/页眉/页脚/日期占位符 + 页码候选（P1-2），不进 title/body。
        - 对标准标题/正文占位符页面与 0.9 行为一致；纯文本框页面为启发式摘要，
          排序稳定、语义一致，不承诺与 0.9 逐字节一致。
        """
        pres = self._require_pres(doc_id)
        pw, ph = _page_size_pt(pres)
        return [
            _summarize_slide(pres.Slides(i), i, pw, ph) for i in range(1, pres.Slides.Count + 1)
        ]

    @readonly_guard
    def read_shapes(
        self,
        slide_idx: int,
        *,
        recursive: bool = True,
        doc_id: str | None = None,
    ) -> list[ShapeInfo]:
        """读取第 slide_idx 页全部 shape 的完整信息（含 group 内后代），不裁剪无文本 shape。

        - shape_id 严格：任何 shape 的 Id 读不到 → ComOperationError（无 0 兜底）。
        - recursive=True 返回 group 容器 + 后代（group_path 外层→内层）；False 仅顶层。
        - coordinate_space：非旋转 group 子元素为 "slide"；旋转 group 内 "unknown"。
        - z_order 为所在集合内 1-based 排序（顶层=ZOrderPosition；group 子元素按兄弟还原）。
        - 颜色 #RRGGBB：仅 solid fill / 可见 line 给色；其余 None（不编假值）。
        - 字体为首 run 语义；空文本 → 全 None。
        """
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        return [
            _record_shape_info(shape, parent_shape_id=pid, group_path=gp, rotated=rot, z_order=z)
            for shape, pid, gp, rot, z in _iter_shape_records(slide.Shapes, recursive=recursive)
        ]

    # --- S1 破坏性编辑 op（读回走 read_shapes） ---
    @destructive
    def set_shape_geometry(
        self,
        slide_idx: int,
        shape_id: int,
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        rotation: float | None = None,
        doc_id: str | None = None,
    ):
        """设置 shape 几何（磅）：left/top/width/height/rotation，只更新传入属性。

        - 至少传一个属性；width/height 必须 > 0；全部数值必须有限（NaN/Inf 拒绝）。
        - group 子元素 left/top 写为**绝对坐标**（探针 #1）；group 包围盒自动重算。
        - 旋转 group 后代：读值被 group 旋转变换不可信（probe #2）→ 传 left/top 抛
          InvalidArgumentError；width/height/rotation 仍可用。
        - 顶层 shape 旋转不影响 Left/Top 读写（probe #2），无特殊顺序要求。
        """
        if all(v is None for v in (left, top, width, height, rotation)):
            raise InvalidArgumentError(
                "set_shape_geometry: 至少提供 left/top/width/height/rotation 之一"
            )
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        located = _find_shape_by_id(slide, shape_id)
        shape = located.shape
        if located.rotated_group_ancestor and (left is not None or top is not None):
            raise InvalidArgumentError(
                "set_shape_geometry: 旋转 group 内后代 left/top 读值不可信"
                "（coordinate_space='unknown'），仅允许 width/height/rotation"
            )
        if left is not None:
            left = _validate_finite_float(left, "left")
        if top is not None:
            top = _validate_finite_float(top, "top")
        if width is not None:
            width = _validate_positive_float(width, "width")
        if height is not None:
            height = _validate_positive_float(height, "height")
        if rotation is not None:
            rotation = _validate_finite_float(rotation, "rotation")
        # 实际 COM 写失败由 @guard_com 统一转 ComOperationError（保留 HRESULT）
        if left is not None:
            shape.Left = left
        if top is not None:
            shape.Top = top
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height
        if rotation is not None:
            shape.Rotation = rotation
        return None

    @destructive
    def set_shape_text(
        self,
        slide_idx: int,
        shape_id: int,
        text: str,
        doc_id: str | None = None,
    ):
        """整体替换 shape 文本（保留样式）。

        - 需要 TextFrame；图片/线条/无文本图形 → InvalidArgumentError。
        - 探针 #4：替换后字体样式完全保留（新单 run 继承原首 run 格式），本 op
          不新增任何样式参数（spec 冻结签名）。
        """
        if not isinstance(text, str):
            raise InvalidArgumentError(
                f"set_shape_text: text 必须是字符串，收到 {type(text).__name__}"
            )
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        located = _find_shape_by_id(slide, shape_id)
        shape = located.shape
        _require_text_frame(shape, "set_shape_text")
        shape.TextFrame.TextRange.Text = text
        return None

    @destructive
    def set_shape_font(
        self,
        slide_idx: int,
        shape_id: int,
        *,
        font_name: str | None = None,
        size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        color: str | None = None,
        doc_id: str | None = None,
    ):
        """设置 shape 字体：font_name/size/bold/italic/color，只更新传入属性。

        - 至少传一个属性；需要 TextFrame。
        - 探针 #3：整范围 TextRange.Font 赋值会传播到**全部 run**，且只赋传入
          属性、其余保留；读取仍走 Runs(1).Font 首 run 语义。
        - color 严格 #RRGGBB；bold/italic 只收真 bool；size > 0。
        """
        if all(v is None for v in (font_name, size, bold, italic, color)):
            raise InvalidArgumentError(
                "set_shape_font: 至少提供 font_name/size/bold/italic/color 之一"
            )
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        located = _find_shape_by_id(slide, shape_id)
        shape = located.shape
        _require_text_frame(shape, "set_shape_font")
        if font_name is not None and not isinstance(font_name, str):
            raise InvalidArgumentError(
                f"set_shape_font: font_name 必须是字符串，收到 {type(font_name).__name__}"
            )
        if size is not None:
            size = _validate_positive_float(size, "size")
        if bold is not None and not isinstance(bold, bool):
            raise InvalidArgumentError(
                f"set_shape_font: bold 必须是 bool，收到 {type(bold).__name__}"
            )
        if italic is not None and not isinstance(italic, bool):
            raise InvalidArgumentError(
                f"set_shape_font: italic 必须是 bool，收到 {type(italic).__name__}"
            )
        if color is not None:
            color = _validate_hex_color(color, "color")
        tr = shape.TextFrame.TextRange
        if font_name is not None:
            tr.Font.Name = font_name
        if size is not None:
            tr.Font.Size = size
        if bold is not None:
            tr.Font.Bold = -1 if bold else 0
        if italic is not None:
            tr.Font.Italic = -1 if italic else 0
        if color is not None:
            tr.Font.Color.RGB = _rgb_to_com(color)
        return None

    @destructive
    def set_shape_fill(
        self,
        slide_idx: int,
        shape_id: int,
        *,
        color: str | None = None,
        transparency: float | None = None,
        doc_id: str | None = None,
    ):
        """设置 shape 填充：color/transparency（0-1），只更新传入属性。

        - color 与 transparency 都未传 → 清除填充（Fill.Visible=0）。
        - 设 color 强制 solid fill（Fill.Solid()）并显示填充；只设 transparency
          时保留原颜色。
        - color 严格 #RRGGBB；transparency [0,1]。
        - shape 不支持 Fill（如线条）→ InvalidArgumentError。
        """
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        located = _find_shape_by_id(slide, shape_id)
        shape = located.shape
        _require_fill_capability(shape, "set_shape_fill")
        if color is None and transparency is None:
            shape.Fill.Visible = 0  # 清除填充
            return None
        if color is not None:
            color = _validate_hex_color(color, "color")
        if transparency is not None:
            transparency = _validate_fraction_0_1(transparency, "transparency")
        fill = shape.Fill
        if color is not None:
            fill.Solid()
            fill.Visible = -1
            fill.ForeColor.RGB = _rgb_to_com(color)
        if transparency is not None:
            fill.Transparency = transparency
        return None

    @destructive
    def set_shape_outline(
        self,
        slide_idx: int,
        shape_id: int,
        *,
        color: str | None = None,
        width: float | None = None,
        visible: bool | None = None,
        doc_id: str | None = None,
    ):
        """设置 shape 轮廓：color/width/visible，只更新传入属性。

        - 至少传一个属性；width > 0；color 严格 #RRGGBB。
        - visible=False 显式隐藏/移除线条；visible=True 显示并尽量保留原色/宽。
        - 设 color 用实线语义（Line.ForeColor.RGB）；width 用磅。
        - shape 不支持 Line → InvalidArgumentError。
        """
        if all(v is None for v in (color, width, visible)):
            raise InvalidArgumentError("set_shape_outline: 至少提供 color/width/visible 之一")
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        located = _find_shape_by_id(slide, shape_id)
        shape = located.shape
        _require_line_capability(shape, "set_shape_outline")
        if color is not None:
            color = _validate_hex_color(color, "color")
        if width is not None:
            width = _validate_positive_float(width, "width")
        if visible is not None and not isinstance(visible, bool):
            raise InvalidArgumentError(
                f"set_shape_outline: visible 必须是 bool，收到 {type(visible).__name__}"
            )
        line = shape.Line
        if color is not None:
            line.ForeColor.RGB = _rgb_to_com(color)
        if width is not None:
            line.Weight = width
        if visible is not None:
            # 最后设可见性：显式 visible 控制最终显示态，色/宽不影响它
            line.Visible = -1 if visible else 0
        return None

    @destructive
    def set_shape_visible(
        self,
        slide_idx: int,
        shape_id: int,
        visible: bool,
        doc_id: str | None = None,
    ):
        """显示/隐藏 shape。

        只收真 bool；映射到 Office 三态常量（-1 显示 / 0 隐藏），与
        read_shapes 的 visible 读取（_shape_visible）一致。
        """
        if not isinstance(visible, bool):
            raise InvalidArgumentError(
                f"set_shape_visible: visible 必须是 bool，收到 {type(visible).__name__}"
            )
        slide = _require_slide(self._require_pres(doc_id), slide_idx)
        located = _find_shape_by_id(slide, shape_id)
        located.shape.Visible = -1 if visible else 0
        return None

    def quit(self, force: bool = False):
        """退出 PowerPoint 会话。

        own 句柄（本库启动的实例）直接退；连到既有 Office 实例默认拒绝
        （不夺走用户正用的窗口），确需退出传 force=True。实例已退（进程
        结束）视为已退出返回 True，不误报失败。
        """
        # 库改全局状态（DisplayAlerts），释放前还原原值
        if not self._owned and not force:
            with suppress(Exception):  # 仅兜底还原，失败不掩盖拒绝语义
                self.app.DisplayAlerts = self._saved_alerts
            raise ComOperationError(
                "连接的是既有 PowerPoint 实例，拒绝退出；确需退出请传 force=True"
            )
        try:
            # P1-3：直接退自持句柄（不重连 ROT 里其它实例），避免误关别人的窗口
            self.app.DisplayAlerts = self._saved_alerts
            pid = core.app_process_pid(self.app, "ppt") or self._pid
            self.app.Quit()
        except Exception as e:  # noqa: BLE001 — com_error/断连异常统一走 liveness 判定
            if not core.doc_alive(self.app):
                return True  # 已退出：liveness 探针证实进程已结束
            raise ComOperationError(f"退出 PowerPoint 失败: {e}") from e
        # Quit 已返回但进程可能残留（RCW/COM server 保持）：按 PID 精确清理
        if pid is not None:
            if not core.wait_process_exit(pid, timeout=2.0):
                core.reap_process(pid)
        else:
            # PID 解析失败（启动早期 ActiveWindow.Hwnd 未就绪）：无法按 PID 清理。
            # liveness 探针兜底——进程仍存活 → 明确失败（force=True 保证退出），
            # 绝不静默 no-op（残留进程会持有文件锁，污染后续覆盖保存）。
            if core.doc_alive(self.app):
                raise ComOperationError(
                    "无法确认 PowerPoint 进程已退出（进程 PID 解析失败）且进程仍存活；"
                    "请手动关闭 PowerPoint 后重试（残留进程会持有文件锁）"
                )
        return None

    def reap_own_process(self) -> None:
        """断连自愈/退出兜底：精确终止本库附着过的实例进程（不碰用户其它实例）。

        server 检测到外部 kill 后重建连接前调用，清掉残留的僵尸实例，
        避免重连附着到半死进程污染后续 op。
        """
        core.reap_process(self._pid)
