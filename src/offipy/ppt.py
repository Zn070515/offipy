"""PowerPoint 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActivePresentation 定位当前文稿。
"""

import math
import os
import shutil
import tempfile
from contextlib import contextmanager, suppress
from typing import Any, NamedTuple

from . import core
from ._comguard import guard_com
from .core import destructive, requires_target
from .exceptions import (
    ComOperationError,
    FileConflictError,
    InvalidArgumentError,
    TargetNotFoundError,
)
from .models import CoordinateSpace, SlideTextRecord, placeholder_type_name
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
            if item.record["shape_id"] not in used and not _is_exempt_text(item.record, pw, ph)
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

    def active_pres(self, doc_id: str | None = None):
        # 显式 doc_id：绑定目标路由，只查文档表；未知/失效句柄抛 TargetNotFoundError。
        # 缺省 active：实时解析 ActivePresentation（doc_id 权威——绝不静默用陈旧的
        # _active_id 快路径，防「用户看到 B、Agent 以为 A」），解析到即并入文档表。
        # P0-8：全程纯探测，绝不隐式 Presentations.Add()。
        if doc_id is not None:
            pres = self._docs.get(doc_id)
            if pres is None or not core.doc_alive(pres):
                raise TargetNotFoundError(
                    f"未知演示文稿句柄: {doc_id!r}（用 list_docs 查看当前打开的）"
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
                f"未知演示文稿句柄: {doc_id!r}（用 list_docs 查看当前打开的）"
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
                pres.SaveAs(dest)
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
            self._require_pres(doc_id).ExportAsFixedFormat(
                dest, FixedFormatType=PP_FIXED_FORMAT_TYPE_PDF, Intent=2, PrintRange=None
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
        pres = self._require_pres(doc_id)
        pres.Slides.Add(pres.Slides.Count + 1, layout)
        return pres.Slides.Count

    @destructive
    def set_title(self, slide_idx: int, text: str, doc_id: str | None = None):
        if not text:
            raise InvalidArgumentError("set_title: text 不能为空")
        slide = self._require_pres(doc_id).Slides(slide_idx)
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
        slide = self._require_pres(doc_id).Slides(slide_idx)
        ph = _placeholder_by_type(slide.Shapes, PP_PLACEHOLDER_BODY)
        if ph is None:
            ph = slide.Shapes.AddTextbox(1, *_BODY_BOX)
        ph.TextFrame.TextRange.Text = "\r".join(lines)
        return ph.Id

    @destructive
    def set_notes(self, slide_idx: int, text: str, doc_id: str | None = None):
        slide = self._require_pres(doc_id).Slides(slide_idx)
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
        slide = self._require_pres(doc_id).Slides(slide_idx)
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
        slide = self._require_pres(doc_id).Slides(slide_idx)
        slide.Shapes.AddPicture(os.path.abspath(path), 0, 0, left, top, width, height)

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
        slide = self._require_pres(doc_id).Slides(slide_idx)
        return [
            item.record
            for item in _collect_text_records(slide, recursive=recursive)
            if include_empty or item.record["text"]
        ]

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
            self.app.Quit()
        except Exception as e:  # noqa: BLE001 — com_error/断连异常统一走 liveness 判定
            if not core.doc_alive(self.app):
                return True  # 已退出：liveness 探针证实进程已结束
            raise ComOperationError(f"退出 PowerPoint 失败: {e}") from e
