"""Word 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActiveDocument 定位当前文档。
"""

import os
from contextlib import contextmanager, suppress
from typing import Any, TypeVar

from . import core
from ._comguard import _COM_ERROR, guard_com
from .core import destructive, readonly_guard, requires_target
from .exceptions import ComOperationError, InvalidArgumentError, TargetNotFoundError
from .paths import default_save_path, ensure_writable

_T = TypeVar("_T")

WD_ALERTS_NONE = 0  # wdAlertsNone：抑制保存/覆盖等模态提示
WD_EXPORT_FORMAT_PDF = 17  # wdExportFormatPDF（ExportAsFixedFormat 的 ExportFormat）
# wdStyleHeading1..3 的 COM 常量值（-2/-3/-4），避免依赖中文/英文样式名
_HEADING_STYLES = {1: -2, 2: -3, 3: -4}


# ===== M5：常量表与解析辅助（Word COM 值按 VBA 常量硬编码，gen_py 实探） =====
# 段落对齐 wdParagraphAlignment
_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
# 下划线 wdUnderline
_UNDERLINE = {"none": 0, "single": 1, "words": 2, "double": 3, "dotted": 4, "wavy": 11}
# 高亮 wdColorIndex（HighlightColorIndex）
_HIGHLIGHT = {
    "none": 0,
    "yellow": 7,
    "green": 11,
    "pink": 5,
    "red": 6,
    "blue": 2,
    "bright_green": 4,
    "turquoise": 3,
}
# 行距规则 wdLineSpacing
_LINE_SPACING = {
    "single": 0,
    "1.5": 1,
    "double": 2,
    "at_least": 3,
    "exactly": 4,
    "multiple": 5,
}
# 页面方向 wdOrientation
_ORIENTATION = {"portrait": 0, "landscape": 1}
# 纸张 wdPaperSize
_PAPER = {"letter": 2, "legal": 4, "a3": 6, "a4": 7, "a5": 9}
# 页码对齐 wdPageNumberAlignment
_PAGE_NUMBER_ALIGN = {"left": 0, "center": 1, "right": 2}
# 查找替换 wdReplace
_REPLACE = {"one": 1, "all": 2}
# 表格线型 wdLineStyle
_LINE_STYLE = {"none": 0, "single": 1, "dot": 2, "double": 7}
# 表格边位置 wdBorderType：1-6 = 上/左/下/右/内横/内竖（Word 枚举为负值，Borders 吃绝对值）
_TABLE_SIDES = {"top": 1, "left": 2, "bottom": 3, "right": 4, "inside-h": 5, "inside-v": 6}
# 行高规则 wdRowHeightRule
_ROW_HEIGHT_RULE = {"auto": 0, "at_least": 1, "exactly": 2}
# 自动调整行为 wdAutoFitBehavior
_AUTOFIT = {"fixed": 0, "content": 1, "window": 2}
# 表格线宽 wdLineWidth（VBA 标准枚举）
_TABLE_LINE_WIDTH = {
    "0.25pt": 2,
    "0.5pt": 4,
    "0.75pt": 6,
    "1pt": 8,
    "1.5pt": 12,
    "2.25pt": 18,
    "3pt": 24,
    "4.5pt": 36,
    "6pt": 48,
}
# 列表样式（add_list 仅支持两种；"bullet" 是 "bulleted" 的旧别名，对外默认仍为 bullet）
_LIST_STYLE = {"numbered": "numbered", "bulleted": "bulleted", "bullet": "bulleted"}


def _rgb(hex_color: str) -> int:
    """把 '#RRGGBB' 转成 COM Long 颜色（Word 与 Excel 同用 COLORREF 公式）。"""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise InvalidArgumentError(f"非法颜色: {hex_color!r}（期望 '#RRGGBB'）")
    try:
        r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        raise InvalidArgumentError(f"非法颜色: {hex_color!r}（期望 '#RRGGBB'）") from None
    return r + (g << 8) + (b << 16)


def _resolve_style(name: str | None, table: dict[str, _T], label: str) -> _T:
    key = (name or "").strip().lower()
    if key not in table:
        raise InvalidArgumentError(f"未知{label}: {name!r}（可选: {', '.join(table)}）")
    return table[key]


def _resolve_table_sides(sides: str | None) -> list[int]:
    """把 all/outside/inside 或逗号分隔的边名解析成 wdBorderType 列表。"""
    name = (sides or "all").strip().lower()
    if name == "all":
        return [1, 2, 3, 4, 5, 6]
    if name == "outside":
        return [1, 2, 3, 4]
    if name == "inside":
        return [5, 6]
    parts = [p.strip() for p in name.split(",") if p.strip()]
    if not parts:
        raise InvalidArgumentError(f"无效表格边: {sides!r}")
    result = []
    for p in parts:
        if p not in _TABLE_SIDES:
            raise InvalidArgumentError(f"未知表格边: {p!r}（可选: {', '.join(_TABLE_SIDES)}）")
        result.append(_TABLE_SIDES[p])
    return result


def _end_range(doc):
    """返回折叠到文档末尾的 Range（用于文末插入）。"""
    rng = doc.Content
    rng.Collapse(0)  # wdCollapseEnd
    return rng


def _check_paragraph(doc, paragraph: int) -> None:
    """段落索引前置校验：1..文档段落数，越界抛 InvalidArgumentError。

    参照 format_text/format_paragraph 的枚举校验，把「集合所要求的成员不存在」
    这类裸 COM 错误前置成 offipy 语义化异常。
    """
    count = doc.Paragraphs.Count
    if paragraph < 1 or paragraph > count:
        raise InvalidArgumentError(f"段落在 1..{count} 范围外（paragraph={paragraph}）")


@guard_com
class WordApp:
    def __init__(self, visible: bool = True, modify_existing_visibility: bool = False):
        self.app, self.created = core.ensure_app(
            "word", visible=visible, modify_existing_visibility=modify_existing_visibility
        )
        # _owned：本库启动的实例才允许 quit() 直接退出；连到既有实例默认拒绝
        self._owned = self.created
        # DisplayAlerts 不再永久静音（P0-5）：按需用 _alerts_scope 临时抑制
        self._saved_alerts = self.app.DisplayAlerts  # quit() 兜底还原
        # 记录本实例进程 PID（断连自愈/退出时精确清理，不误杀用户其它实例）
        self._pid = core.app_process_pid(self.app, "word")
        self._docs: dict[str, Any] = {}  # doc_id → 文档句柄（P2-2 多文档）
        self._active_id: str | None = None
        self._seq = 0

    @contextmanager
    def _alerts_scope(self, value: int = WD_ALERTS_NONE):
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
            for did, doc in self._docs.items():
                if self._stable_identity(doc) == ident:
                    self._docs[did] = obj  # 复用 doc_id，换用实时句柄
                    self._active_id = did
                    return did
        self._seq += 1
        did = f"doc{self._seq}"
        self._docs[did] = obj
        self._active_id = did
        return did

    def _sync_registered(self, obj) -> str:
        """把实时解析到的句柄并入文档表：已登记则复用并置活动，否则登记为新文档。"""
        for did, doc in self._docs.items():
            if doc is obj:
                self._active_id = did
                return did
        return self._register(obj)

    # --- 文档（P2-2 多文档：doc_id 显式路由，缺省走活动） ---
    def new_doc(self) -> str:
        """新建空白文档，登记进文档表并设为活动。返回 doc_id。"""
        return self._register(self.app.Documents.Add())

    def open_doc(self, path: str) -> str:
        """打开现有文档并设为活动。返回 doc_id。"""
        if not os.path.isfile(path):
            raise InvalidArgumentError(f"源文件不存在: {path}")
        return self._register(self.app.Documents.Open(path))

    def active_doc(self, doc_id: str | None = None):
        # 显式 doc_id：绑定目标路由，只查文档表；未知/失效句柄抛 TargetNotFoundError。
        # 缺省 active：实时解析 ActiveDocument（doc_id 权威——绝不静默用陈旧的
        # _active_id 快路径，防「用户看到 B、Agent 以为 A」），解析到即并入文档表。
        # P0-8：全程纯探测，绝不隐式 Documents.Add()。
        if doc_id is not None:
            doc = self._docs.get(doc_id)
            if doc is None or not core.doc_alive(doc):
                raise TargetNotFoundError(
                    f"未知文档句柄: {doc_id!r}（当前会话未打开；doc_id 只在同会话内有效，"
                    f"本地直连 Word() 与会话式 Remote*/CLI/HTTP 互不相通；用本会话 list_docs 核对）"
                )
            return doc
        doc = core.active_doc("word", "ActiveDocument")
        if doc is not None:
            self._sync_registered(doc)
            return doc
        doc = self.app.ActiveDocument
        if doc is None:
            return None
        self._sync_registered(doc)
        return doc

    def _require_doc(self, doc_id: str | None = None):
        """操作前置：目标文档不存在则抛 TargetNotFoundError，不隐式创建。"""
        doc = self.active_doc(doc_id)
        if doc is None:
            raise TargetNotFoundError("没有打开的 Word 文档，请先 new_doc/open_doc")
        return doc

    def activate(self, doc_id: str) -> str:
        """把指定文档设为活动目标并同步真实 UI（Document.Activate）。

        未知句柄抛 TargetNotFoundError。
        """
        doc = self._docs.get(doc_id)
        if doc is None or not core.doc_alive(doc):
            raise TargetNotFoundError(
                f"未知文档句柄: {doc_id!r}（当前会话未打开；doc_id 只在同会话内有效，"
                f"本地直连 Word() 与会话式 Remote*/CLI/HTTP 互不相通；用本会话 list_docs 核对）"
            )
        old = self._active_id
        self._active_id = doc_id
        try:
            doc.Activate()  # 同步 Word 真实活动文档，防止用户焦点漂移
        except Exception as e:
            self._active_id = old  # 同步不上则回滚，不静默假活
            raise ComOperationError(f"激活文档 {doc_id} 失败: {e}") from e
        return doc_id

    def list_docs(self) -> dict:
        """当前打开的文档表：{doc_id: {"name", "path", "active"}}。只报已登记句柄，不隐式枚举。

        P1-5：先并入真实活动焦点（active_doc 解析 ActiveDocument 进文档表、
        刷新 _active_id），active 标记跟随用户当前看到的文档，不报陈旧焦点。
        """
        with suppress(Exception):
            self.active_doc()
        out = {}
        for did, doc in self._docs.items():
            if not core.doc_alive(doc):
                continue
            try:
                name = doc.Name
            except Exception:
                name = None
            try:
                path = doc.FullName
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
            doc = self.active_doc(doc_id)
            resolved = doc_id
        else:
            doc = self.active_doc()
            if doc is None:
                return None
            active_id = self._active_id
            assert active_id is not None  # active_doc 非 None 时活动 id 必已同步
            resolved = active_id
        try:
            name = doc.Name
        except Exception:
            name = None
        try:
            path = doc.FullName
        except Exception:
            path = None
        return {"app": "word", "doc_id": resolved, "name": name, "path": path}

    @destructive
    def close_doc(self, save: bool = True, doc_id: str | None = None):
        """关闭文档（doc_id 缺省为活动）。

        save=True → 先保存（从未保存过则自动落盘用户数据目录，不弹另存为）并返回
        保存路径；save=False → 直接关闭不保存、不弹对话框，返回 None。
        """
        doc = self._require_doc(doc_id)
        did = doc_id if doc_id is not None else self._active_id
        if save:
            path = doc.FullName if doc.Path else self.save(doc_id=did)
            with self._alerts_scope():
                doc.Close(SaveChanges=-1)
        else:
            path = None
            with self._alerts_scope():
                doc.Saved = True  # 兜底：确保 Close 不触发保存提示
                doc.Close(SaveChanges=0)
        if did is not None:
            self._docs.pop(did, None)
            if self._active_id == did:
                self._active_id = None
        return path

    @destructive
    def save(self, path: str | None = None, overwrite: bool = False, doc_id: str | None = None):
        """保存文档并返回绝对路径。

        给 path → 另存到该路径；未给 path → 已保存过的存回原路径，从未保存过的
        自动落盘 <用户数据目录>/documents/<名字>_<时间戳>.docx（不弹另存为对话框）。
        """
        if path:
            dest = ensure_writable(path, overwrite)  # 覆盖保护先于触 COM（fail-fast）
            doc = self._require_doc(doc_id)
            with self._alerts_scope():
                doc.SaveAs2(dest)
            return dest
        doc = self._require_doc(doc_id)
        with self._alerts_scope():
            if doc.Path:  # 已有保存路径 → 原位保存
                doc.Save()
                return doc.FullName
            dest = default_save_path(doc.Name, ".docx")
            doc.SaveAs2(dest)
            return dest

    @requires_target
    def save_pdf(self, path: str, overwrite: bool = False, doc_id: str | None = None):
        dest = ensure_writable(path, overwrite)
        with self._alerts_scope():
            self._require_doc(doc_id).ExportAsFixedFormat(dest, ExportFormat=WD_EXPORT_FORMAT_PDF)

    # --- 内容 ---
    @destructive
    def write(self, text: str, doc_id: str | None = None):
        self._require_doc(doc_id).Content.InsertAfter(text)

    @destructive
    def write_line(self, text: str, doc_id: str | None = None):
        self._require_doc(doc_id).Content.InsertAfter(text + "\r\n")

    @destructive
    def add_heading(self, text: str, level: int = 1, doc_id: str | None = None):
        self.write_line(text, doc_id)
        doc = self._require_doc(doc_id)
        style = _HEADING_STYLES.get(level, -2)
        # write_line = Content.InsertAfter(text + "\r\n")：文本落在末尾空段之前的 Count-1 段
        # （Count 是空尾段，实机验证）。若给 Count 上样式，随后正文会继承标题样式，
        # 目录（按标题样式收集）为空。
        doc.Paragraphs(doc.Paragraphs.Count - 1).Style = style

    @destructive
    def add_table(self, rows: int, cols: int, doc_id: str | None = None):
        if rows < 1 or cols < 1:
            raise InvalidArgumentError(f"add_table: 行列必须 ≥ 1（rows={rows}, cols={cols}）")
        doc = self._require_doc(doc_id)
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd：折叠到文末
        doc.Tables.Add(rng, rows, cols)
        return doc.Tables.Count

    @destructive
    def set_table_cell(
        self, table_idx: int, row: int, col: int, text: str, doc_id: str | None = None
    ):
        self._require_doc(doc_id).Tables(table_idx).Cell(row, col).Range.Text = text

    # --- 样式系统：文字格式 ---
    @destructive
    def format_text(
        self,
        paragraph: int,
        bold: bool | None = None,
        italic: bool | None = None,
        size: float | None = None,
        name: str | None = None,
        color: str | None = None,
        underline: str | None = None,
        highlight: str | None = None,
        doc_id: str | None = None,
    ):
        doc = self._require_doc(doc_id)
        _check_paragraph(doc, paragraph)
        font = doc.Paragraphs(paragraph).Range.Font
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if size is not None:
            font.Size = size
        if name is not None:
            font.Name = name
        if color is not None:
            font.Color = _rgb(color)
        if underline is not None:
            font.Underline = _resolve_style(underline, _UNDERLINE, "下划线")
        if highlight is not None:
            font.HighlightColorIndex = _resolve_style(highlight, _HIGHLIGHT, "高亮色")

    # --- 样式系统：段落格式 ---
    @destructive
    def format_paragraph(
        self,
        paragraph: int,
        alignment: str | None = None,
        line_spacing: str | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        left_indent: float | None = None,
        first_line_indent: float | None = None,
        doc_id: str | None = None,
    ):
        doc = self._require_doc(doc_id)
        _check_paragraph(doc, paragraph)
        fmt = doc.Paragraphs(paragraph).Format
        if alignment is not None:
            fmt.Alignment = _resolve_style(alignment, _ALIGN, "对齐")
        if line_spacing is not None:
            fmt.LineSpacingRule = _resolve_style(line_spacing, _LINE_SPACING, "行距")
        if space_before is not None:
            fmt.SpaceBefore = space_before
        if space_after is not None:
            fmt.SpaceAfter = space_after
        if left_indent is not None:
            fmt.LeftIndent = left_indent
        if first_line_indent is not None:
            fmt.FirstLineIndent = first_line_indent

    # --- 页面结构：页眉页脚 / 页码 / 页面设置 ---
    @destructive
    def set_header_text(self, text: str, section: int = 1, doc_id: str | None = None):
        self._require_doc(doc_id).Sections(section).Headers(1).Range.Text = text

    @destructive
    def set_footer_text(self, text: str, section: int = 1, doc_id: str | None = None):
        self._require_doc(doc_id).Sections(section).Footers(1).Range.Text = text

    @destructive
    def add_page_number(
        self,
        alignment: str = "right",
        color: str | None = None,
        size: float | None = None,
        doc_id: str | None = None,
    ):
        hf = self._require_doc(doc_id).Sections(1).Footers(1)
        hf.Range.Text = ""  # 清空页脚，避免与既有文本叠加
        # PageNumber 对象没有 Range 属性（gen_py 实测 AttributeError），
        # 样式与文本都落在页脚 Range 上（页码域在其中）。
        hf.PageNumbers.Add(
            PageNumberAlignment=_resolve_style(alignment, _PAGE_NUMBER_ALIGN, "页码对齐")
        )
        if color is not None:
            hf.Range.Font.Color = _rgb(color)
        if size is not None:
            hf.Range.Font.Size = size
        return hf.Range.Text

    @destructive
    def page_setup(
        self,
        orientation: str | None = None,
        paper: str | None = None,
        left_margin: float | None = None,
        right_margin: float | None = None,
        top_margin: float | None = None,
        bottom_margin: float | None = None,
        gutter: float | None = None,
        doc_id: str | None = None,
    ):
        ps = self._require_doc(doc_id).PageSetup
        if orientation is not None:
            ps.Orientation = _resolve_style(orientation, _ORIENTATION, "页面方向")
        if paper is not None:
            ps.PaperSize = _resolve_style(paper, _PAPER, "纸张")
        if left_margin is not None:
            ps.LeftMargin = left_margin
        if right_margin is not None:
            ps.RightMargin = right_margin
        if top_margin is not None:
            ps.TopMargin = top_margin
        if bottom_margin is not None:
            ps.BottomMargin = bottom_margin
        if gutter is not None:
            ps.Gutter = gutter

    # --- 页面结构：目录 ---
    @destructive
    def insert_toc(self, levels: int = 3, doc_id: str | None = None):
        doc = self._require_doc(doc_id)
        doc.TablesOfContents.Add(
            doc.Range(0, 0), UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=levels
        )
        return doc.TablesOfContents.Count

    @destructive
    def update_toc(self, doc_id: str | None = None):
        doc = self._require_doc(doc_id)
        doc.TablesOfContents(1).Update()
        return doc.TablesOfContents.Count

    # --- 列表 ---
    @destructive
    def add_list(self, lines: list[str], style: str = "bullet", doc_id: str | None = None):
        if not lines:
            raise InvalidArgumentError("add_list: lines 不能为空")
        list_kind = _resolve_style(style, _LIST_STYLE, "列表样式")
        doc = self._require_doc(doc_id)
        start = doc.Paragraphs.Count + 1  # 第一个新段落的序号
        for line in lines:
            self.write_line(line, doc_id)
        end = doc.Paragraphs.Count
        # Range 起点前移一个字符，落在上一段段末标记上，
        # 避免 ApplyBulletDefault 跳过 range 首段（Word 段落边界行为）
        rng = doc.Range(doc.Paragraphs(start).Range.Start - 1, doc.Paragraphs(end).Range.End)
        if list_kind == "numbered":
            rng.ListFormat.ApplyNumberDefault()
        else:
            rng.ListFormat.ApplyBulletDefault()
        return len(lines)

    # --- 表格增强 ---
    @destructive
    def merge_table_cells(
        self,
        table_idx: int,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        doc_id: str | None = None,
    ):
        t = self._require_doc(doc_id).Tables(table_idx)
        t.Cell(start_row, start_col).Merge(t.Cell(end_row, end_col))

    @destructive
    def set_table_border(
        self,
        table_idx: int,
        style: str = "single",
        weight: str | None = None,
        color: str | None = None,
        sides: str | None = None,
        doc_id: str | None = None,
    ):
        t = self._require_doc(doc_id).Tables(table_idx)
        const = _resolve_style(style, _LINE_STYLE, "线型")
        for idx in _resolve_table_sides(sides):
            b = t.Borders(idx)
            b.LineStyle = const
            if weight is not None:
                b.LineWidth = _resolve_style(weight, _TABLE_LINE_WIDTH, "线宽")
            if color is not None:
                b.Color = _rgb(color)

    @destructive
    def set_table_col_width(
        self, table_idx: int, col: int, width: float, doc_id: str | None = None
    ):
        if width < 7:
            raise InvalidArgumentError(f"set_table_col_width: 列宽需 ≥ 7pt（width={width}）")
        table = self._require_doc(doc_id).Tables(table_idx)
        # autofit 后 AllowAutoFit 仍约束列宽赋值（Word 抛「数值超出范围」），先复位再设宽
        table.AllowAutoFit = False
        try:
            table.Columns(col).Width = width
        except _COM_ERROR:
            # 合并单元格后 Columns(col) 本身就被拒访（「表格有混合的单元格宽度」），
            # 连 Columns(col).Cells 也一样。彻底绕开列对象：逐行取 table.Cell(r, col)，
            # 每个 cell 有独立 Width 可写，先合并后调列宽也能成立。
            for r in range(1, table.Rows.Count + 1):
                try:
                    table.Cell(r, col).Width = width
                except _COM_ERROR:
                    # 目标列落在合并区域内（Cell(r, col) 不独立存在），无法独立设宽——
                    # 语义化转义，而非漏原始 COM 消息。
                    raise ComOperationError(
                        f"set_table_col_width: 目标列 {col} 在第 {r} 行被合并区域覆盖，"
                        "无法独立设宽；请对合并后的整体单元格设宽或先拆分合并"
                    ) from None

    @destructive
    def set_table_row_height(
        self,
        table_idx: int,
        row: int,
        height: float,
        rule: str = "at_least",
        doc_id: str | None = None,
    ):
        r = self._require_doc(doc_id).Tables(table_idx).Rows(row)
        r.Height = height
        r.HeightRule = _resolve_style(rule, _ROW_HEIGHT_RULE, "行高规则")

    @destructive
    def autofit_table(self, table_idx: int, behavior: str = "content", doc_id: str | None = None):
        self._require_doc(doc_id).Tables(table_idx).AutoFitBehavior(
            _resolve_style(behavior, _AUTOFIT, "自动调整")
        )

    # --- 文档辅助 ---
    @destructive
    def find_replace(
        self,
        find: str,
        replace: str,
        match_case: bool = False,
        whole_word: bool = False,
        replace_all: bool = True,
        doc_id: str | None = None,
    ):
        f = self._require_doc(doc_id).Content.Find
        f.Execute(
            FindText=find,
            ReplaceWith=replace,
            Replace=_REPLACE["all" if replace_all else "one"],
            Forward=True,
            MatchCase=match_case,
            MatchWholeWord=whole_word,
        )

    @destructive
    def insert_image(
        self,
        path: str,
        width: float | None = None,
        height: float | None = None,
        doc_id: str | None = None,
    ):
        if not os.path.isfile(path):
            raise InvalidArgumentError(f"源文件不存在: {path}")
        doc = self._require_doc(doc_id)
        shape = doc.InlineShapes.AddPicture(
            os.path.normpath(os.path.abspath(path)), Range=_end_range(doc)
        )
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height
        return doc.InlineShapes.Count

    @destructive
    def insert_page_break(self, doc_id: str | None = None):
        _end_range(self._require_doc(doc_id)).InsertBreak(7)  # wdPageBreak

    # --- 只读辅助（支撑 Agent 文本层读回迭代） ---
    @readonly_guard
    def read_doc_text(self, doc_id: str | None = None):
        """读取当前文档全文文本（只读，不改任何状态）。

        表格 cell 结束符 \\x07 替换成 " | "、段落符 \\r\\n / \\r 归一化成 \\n
        （\\r\\n 先行避免双换行），消费端直接得到可读结构，无需自行剥离 Word 原始标记。
        """
        text = self._require_doc(doc_id).Content.Text
        return text.replace("\x07", " | ").replace("\r\n", "\n").replace("\r", "\n")

    def quit(self, force: bool = False):
        """退出 Word 会话。

        own 句柄（本库启动的实例）直接退；连到既有 Office 实例默认拒绝
        （不夺走用户正用的窗口），确需退出传 force=True。实例已退（进程
        结束）视为已退出返回 True，不误报失败。
        """
        # 库改全局状态（DisplayAlerts），释放前还原原值
        if not self._owned and not force:
            with suppress(Exception):  # 仅兜底还原，失败不掩盖拒绝语义
                self.app.DisplayAlerts = self._saved_alerts
            raise ComOperationError("连接的是既有 Word 实例，拒绝退出；确需退出请传 force=True")
        try:
            # P1-3：直接退自持句柄（不重连 ROT 里其它实例），避免误关别人的窗口
            self.app.DisplayAlerts = self._saved_alerts
            pid = core.app_process_pid(self.app, "word") or self._pid
            self.app.Quit()
        except Exception as e:  # noqa: BLE001 — com_error/断连异常统一走 liveness 判定
            if not core.doc_alive(self.app):
                return True  # 已退出：liveness 探针证实进程已结束
            raise ComOperationError(f"退出 Word 失败: {e}") from e
        # Quit 已返回但进程可能残留（RCW/COM server 保持）：按 PID 精确清理
        if not core.wait_process_exit(pid, timeout=2.0):
            core.reap_process(pid)
        return None

    def reap_own_process(self) -> None:
        """断连自愈/退出兜底：精确终止本库附着过的实例进程（不碰用户其它实例）。

        server 检测到外部 kill 后重建连接前调用，清掉残留的僵尸实例，
        避免重连附着到半死进程污染后续 op。
        """
        core.reap_process(self._pid)
