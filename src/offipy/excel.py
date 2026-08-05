"""Excel 会话式自动化。

基于 core 的会话管理：每次调用重连同一个 Excel 实例，跨进程时通过
ActiveWorkbook 定位当前工作簿（即用户在 Excel 里当前激活的那个）。
"""

import re
from contextlib import contextmanager, suppress
from typing import Any

from . import core
from ._comguard import _COM_ERROR, guard_com
from .core import destructive
from .exceptions import ComOperationError, InvalidArgumentError, TargetNotFoundError
from .paths import default_save_path, ensure_writable

# ExportAsFixedFormat 的类型常量
XL_TYPE_PDF = 0


_CELL_RE = re.compile(r"^([A-Za-z]{1,3})(\d{1,7})$")
# Excel 真实 grid 上限：XFD 列（16384）× 1048576 行
_MAX_COL = 16384
_MAX_ROW = 1048576


def _parse_cell(cell: str):
    r"""把 'A1' 解析成 (row, col)，行列为 1 基。

    收严为 Excel 真实坐标：`^([A-Za-z]{1,3})(\d{1,7})$`。畸形（如 'A1B2'——
    不再被字母/数字拆分误读）或越界（列 > XFD、行 > 1048576）抛
    InvalidArgumentError。
    """
    m = _CELL_RE.match(str(cell))
    if m is None:
        raise InvalidArgumentError(f"非法单元格: {cell!r}（期望如 'A1'，列 ≤ XFD）")
    col_part, row_part = m.groups()
    col = 0
    for ch in col_part.upper():
        col = col * 26 + (ord(ch) - ord("A") + 1)
    row = int(row_part)
    if row < 1 or col > _MAX_COL or row > _MAX_ROW:
        raise InvalidArgumentError(f"单元格越界: {cell!r}（上限 XFD1048576）")
    return row, col


def _rgb(hex_color: str) -> int:
    """把 '#RRGGBB' 转成 Excel 的 BGR 整数颜色；非法颜色抛 InvalidArgumentError。"""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise InvalidArgumentError(f"非法颜色: {hex_color!r}（期望 '#RRGGBB'）")
    try:
        r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        raise InvalidArgumentError(f"非法颜色: {hex_color!r}（期望 '#RRGGBB'）") from None
    return r + (g << 8) + (b << 16)


def _normalize_range(values):
    """把 COM Range.Value 归一成二维 list：单格→[[v]]，空→[]。

    COM 按行外层/列内层返回：A1:B2 → ((r1c1,r1c2),(r2c1,r2c2))；单格返回
    标量本身（部分版本返回 1x1 元组）；空区域返回 None。
    """
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        if isinstance(values[0], (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


# ===== M4：常量表与解析辅助（Excel COM 值按 VBA 常量硬编码） =====
# 边框位置 BorderIndex：7/8/9/10=左/上/下/右，11/12=内竖/内横
_BORDER_INDEX = {
    "left": 7,
    "top": 8,
    "bottom": 9,
    "right": 10,
    "inside-v": 11,
    "inside-h": 12,
}
# 线型 xlLineStyle
_LINE_STYLE = {
    "continuous": 1,
    "dash": -4115,
    "dash-dot": 4,
    "dash-dot-dot": 5,
    "dot": -4118,
    "double": -4119,
    "none": -4142,
    "slant-dash-dot": 13,
}
# 线宽 xlBorderWeight
_BORDER_WEIGHT = {"hairline": 1, "thin": 2, "medium": -4138, "thick": 4}
# 条件格式运算符 xlFormatConditionOperator
_COND_OPERATOR = {
    "between": 1,
    "not_between": 2,
    "equal": 3,
    "not_equal": 4,
    "greater": 5,
    "less": 6,
    "greater_equal": 7,
    "less_equal": 8,
}
# 页面方向 xlPageOrientation
_ORIENTATION = {"portrait": 1, "landscape": 2}
# 纸张 xlPaperSize
_PAPER_SIZE = {"letter": 1, "a3": 8, "a4": 9}


def _resolve_sides(side: str | None) -> list[int]:
    """把 all/outside/inside 或逗号分隔的具体边名解析成 BorderIndex 列表。"""
    name = (side or "all").strip().lower()
    if name == "all":
        return [7, 8, 9, 10, 11, 12]
    if name == "outside":
        return [7, 8, 9, 10]
    if name == "inside":
        return [11, 12]
    parts = [p.strip() for p in name.split(",") if p.strip()]
    if not parts:
        raise InvalidArgumentError(f"无效边框侧: {side!r}")
    result = []
    for p in parts:
        if p not in _BORDER_INDEX:
            raise InvalidArgumentError(f"未知边框侧: {p!r}（可选: {', '.join(_BORDER_INDEX)}）")
        result.append(_BORDER_INDEX[p])
    return result


def _resolve_style(name: str | None, table: dict[str, int], label: str) -> int:
    key = (name or "").strip().lower()
    if key not in table:
        raise InvalidArgumentError(f"未知{label}: {name!r}（可选: {', '.join(table)}）")
    return table[key]


@guard_com
class ExcelApp:
    def __init__(self, visible: bool = True, modify_existing_visibility: bool = False):
        self.app, self.created = core.ensure_app(
            "excel", visible=visible, modify_existing_visibility=modify_existing_visibility
        )
        # _owned：本库启动的实例才允许 quit() 直接退出；连到既有实例默认拒绝
        self._owned = self.created
        # DisplayAlerts 不再永久静音（P0-5）：按需用 _alerts_scope 临时抑制
        self._saved_alerts = self.app.DisplayAlerts  # quit() 兜底还原
        self._docs: dict[str, Any] = {}  # doc_id → 工作簿句柄（P2-2 多文档）
        self._active_id: str | None = None
        self._seq = 0

    @contextmanager
    def _alerts_scope(self, value: bool = False):
        """临时抑制模态对话框；退出时（含异常路径）还原 DisplayAlerts 原值。"""
        prev = self.app.DisplayAlerts
        self.app.DisplayAlerts = value
        try:
            yield
        finally:
            self.app.DisplayAlerts = prev

    def _stable_identity(self, obj):
        """稳定身份键（P0-4）：已保存 → (FullName.lower(), None)；未保存 → (None, Name.lower())。

        pywin32 的 wrapper 每次获取都可能是新对象（`is` 不成立），但底层文档的
        FullName/Name 稳定——同文件重开/重连据此复用同一 doc_id。双 None 表示
        无法识别，跳过匹配（防死句柄误复用）。
        """
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
            for did, book in self._docs.items():
                if self._stable_identity(book) == ident:
                    self._docs[did] = obj  # 复用 doc_id，换用实时句柄
                    self._active_id = did
                    return did
        self._seq += 1
        did = f"book{self._seq}"
        self._docs[did] = obj
        self._active_id = did
        return did

    def _sync_registered(self, obj) -> str:
        """把实时解析到的句柄并入文档表：已登记则复用并置活动，否则登记为新文档。"""
        for did, book in self._docs.items():
            if book is obj:
                self._active_id = did
                return did
        return self._register(obj)

    # --- 工作簿（P2-2 多文档：doc_id 显式路由，缺省走活动） ---
    def new_book(self) -> str:
        """新建空白工作簿，登记进文档表并设为活动。返回 doc_id。"""
        return self._register(self.app.Workbooks.Add())

    def open_book(self, path: str) -> str:
        """打开现有工作簿并设为活动。返回 doc_id。"""
        return self._register(self.app.Workbooks.Open(path))

    def active_book(self, doc_id: str | None = None):
        # 显式 doc_id：绑定目标路由，只查文档表；未知/失效句柄抛 TargetNotFoundError。
        # 缺省 active：实时解析 ActiveWorkbook（doc_id 权威——绝不静默用陈旧的
        # _active_id 快路径，防「用户看到 B、Agent 以为 A」），解析到即并入文档表。
        # P0-8：全程纯探测，绝不隐式 Workbooks.Add()。
        if doc_id is not None:
            book = self._docs.get(doc_id)
            if book is None or not core.doc_alive(book):
                raise TargetNotFoundError(
                    f"未知工作簿句柄: {doc_id!r}（用 list_docs 查看当前打开的）"
                )
            return book
        book = core.active_doc("excel", "ActiveWorkbook")
        if book is not None:
            self._sync_registered(book)
            return book
        book = self.app.ActiveWorkbook
        if book is None:
            return None
        self._sync_registered(book)
        return book

    def _require_book(self, doc_id: str | None = None):
        """操作前置：目标工作簿不存在则抛 TargetNotFoundError，不隐式创建。"""
        book = self.active_book(doc_id)
        if book is None:
            raise TargetNotFoundError("没有打开的工作簿，请先 new_book/open_book")
        return book

    def activate(self, doc_id: str) -> str:
        """把指定文档设为活动目标并同步真实 UI（Workbook.Activate）。

        未知句柄抛 TargetNotFoundError。
        """
        book = self._docs.get(doc_id)
        if book is None or not core.doc_alive(book):
            raise TargetNotFoundError(f"未知工作簿句柄: {doc_id!r}（用 list_docs 查看当前打开的）")
        old = self._active_id
        self._active_id = doc_id
        try:
            book.Activate()  # 同步 Excel 真实活动工作簿，防止用户焦点漂移
        except Exception as e:
            self._active_id = old  # 同步不上则回滚，不静默假活
            raise ComOperationError(f"激活工作簿 {doc_id} 失败: {e}") from e
        return doc_id

    def list_docs(self) -> dict:
        """当前打开的文档表：{doc_id: {"name", "path", "active"}}。只报已登记句柄，不隐式枚举。"""
        out = {}
        for did, book in self._docs.items():
            if not core.doc_alive(book):
                continue
            try:
                name = book.Name
            except Exception:
                name = None
            try:
                path = book.FullName
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
            book = self.active_book(doc_id)
            resolved = doc_id
        else:
            book = self.active_book()
            if book is None:
                return None
            active_id = self._active_id
            assert active_id is not None  # active_book 非 None 时活动 id 必已同步
            resolved = active_id
        try:
            name = book.Name
        except Exception:
            name = None
        try:
            path = book.FullName
        except Exception:
            path = None
        return {"app": "excel", "doc_id": resolved, "name": name, "path": path}

    @destructive
    def close_book(self, save: bool = True, doc_id: str | None = None):
        """关闭工作簿（doc_id 缺省为活动）。

        save=True → 先保存（从未保存过则自动落盘用户数据目录，不弹另存为）并返回
        保存路径；save=False → 直接关闭不保存、不弹对话框，返回 None。
        """
        book = self._require_book(doc_id)
        did = doc_id if doc_id is not None else self._active_id
        if save:
            path = book.FullName if book.Path else self.save(doc_id=did)
            with self._alerts_scope():
                book.Close(SaveChanges=1)
        else:
            path = None
            with self._alerts_scope():
                # Excel 对「从未保存过的脏工作簿」Close 会弹另存为，即使
                # SaveChanges=xlDoNotSaveChanges(2)+DisplayAlerts=False 也一样；
                # 先标 Saved=True 让 Excel 认为无未保存更改，才能根治不弹窗。
                book.Saved = True
                book.Close(SaveChanges=2)
        if did is not None:
            self._docs.pop(did, None)
            if self._active_id == did:
                self._active_id = None
        return path

    @destructive
    def save(self, path: str | None = None, overwrite: bool = False, doc_id: str | None = None):
        """保存工作簿并返回绝对路径。

        给 path → 另存到该路径；未给 path → 已保存过的存回原路径，从未保存过的
        自动落盘 <用户数据目录>/documents/<名字>_<时间戳>.xlsx（不弹另存为对话框）。
        """
        if path:
            dest = ensure_writable(path, overwrite)  # 覆盖保护先于触 COM（fail-fast）
            book = self._require_book(doc_id)
            with self._alerts_scope():
                # COM 的 SaveAs 不认正斜杠，必须规范为反斜杠绝对路径
                book.SaveAs(dest)
            return dest
        book = self._require_book(doc_id)
        with self._alerts_scope():
            if book.Path:  # 已有保存路径 → 原位保存
                book.Save()
                return book.FullName
            dest = default_save_path(book.Name, ".xlsx")
            book.SaveAs(dest)
            return dest

    def save_pdf(self, path: str, overwrite: bool = False, doc_id: str | None = None):
        dest = ensure_writable(path, overwrite)
        with self._alerts_scope():
            self._require_book(doc_id).ExportAsFixedFormat(XL_TYPE_PDF, dest)

    # --- 工作表 ---
    def _ws(self, sheet, doc_id: str | None = None):
        book = self._require_book(doc_id)
        if isinstance(sheet, str):
            try:
                return book.Worksheets(sheet)
            except _COM_ERROR:
                if not sheet.isdigit():
                    raise ComOperationError(f"工作表不存在: {sheet!r}") from None
                try:
                    return book.Worksheets(int(sheet))
                except (ValueError, _COM_ERROR):
                    raise ComOperationError(f"工作表不存在: {sheet!r}") from None
        return book.Worksheets(sheet)

    @destructive
    def add_sheet(self, name: str, doc_id: str | None = None):
        book = self._require_book(doc_id)
        ws = book.Worksheets.Add()
        ws.Name = name
        return ws

    # --- 单元格 ---
    @destructive
    def set_cell(self, sheet, cell: str, value, doc_id: str | None = None):
        row, col = _parse_cell(cell)
        self._ws(sheet, doc_id).Cells(row, col).Value = value

    def get_cell(self, sheet, cell: str, doc_id: str | None = None):
        row, col = _parse_cell(cell)
        return self._ws(sheet, doc_id).Cells(row, col).Value

    @destructive
    def set_range(self, sheet, range_addr: str, values, doc_id: str | None = None):
        self._ws(sheet, doc_id).Range(range_addr).Value = values

    @destructive
    def set_col_width(self, sheet, col, width, doc_id: str | None = None):
        self._ws(sheet, doc_id).Columns(col).ColumnWidth = width

    def read_range(self, sheet, range_addr, doc_id: str | None = None):
        """读取区域值，返回二维 list（行→列）。只读，不改状态。"""
        return _normalize_range(self._ws(sheet, doc_id).Range(range_addr).Value)

    # --- 格式化 ---
    @destructive
    def format_cell(
        self,
        sheet,
        cell: str,
        bold=None,
        size=None,
        italic=None,
        bg=None,
        fg=None,
        align=None,
        doc_id: str | None = None,
    ):
        row, col = _parse_cell(cell)
        cell_obj = self._ws(sheet, doc_id).Cells(row, col)
        font = cell_obj.Font
        if bold is not None:
            font.Bold = bold
        if size is not None:
            font.Size = size
        if italic is not None:
            font.Italic = italic
        if bg is not None:
            cell_obj.Interior.Color = _rgb(bg)
        if fg is not None:
            font.Color = _rgb(fg)
        if align is not None:
            cell_obj.HorizontalAlignment = align

    # --- 合并单元格 ---
    @destructive
    def merge_cells(self, sheet, range_addr: str, doc_id: str | None = None):
        self._ws(sheet, doc_id).Range(range_addr).Merge()

    @destructive
    def unmerge_cells(self, sheet, range_addr: str, doc_id: str | None = None):
        self._ws(sheet, doc_id).Range(range_addr).UnMerge()

    # --- 边框 ---
    @destructive
    def set_border(
        self,
        sheet,
        range_addr: str,
        side: str = "all",
        style: str = "continuous",
        weight: str = "thin",
        color: str | None = None,
        doc_id: str | None = None,
    ):
        ws = self._ws(sheet, doc_id)
        rng = ws.Range(range_addr)
        style_const = _resolve_style(style, _LINE_STYLE, "线型")
        weight_const = _resolve_style(weight, _BORDER_WEIGHT, "线宽")
        for idx in _resolve_sides(side):
            b = rng.Borders(idx)
            b.LineStyle = style_const
            b.Weight = weight_const
            if color is not None:
                b.Color = _rgb(color)

    # --- 冻结窗格 ---
    @destructive
    def freeze_panes(self, sheet, rows: int = 0, cols: int = 0, doc_id: str | None = None):
        if rows < 0 or cols < 0:
            raise InvalidArgumentError(f"rows/cols 必须 ≥0，收到 rows={rows}, cols={cols}")
        ws = self._ws(sheet, doc_id)
        ws.Activate()
        if rows == 0 and cols == 0:
            self.app.ActiveWindow.FreezePanes = False
        else:
            # 冻结 cell 左上方区域：选中 (rows+1, cols+1) 再开 FreezePanes
            ws.Cells(rows + 1, cols + 1).Select()
            self.app.ActiveWindow.FreezePanes = True

    # --- 打印设置 ---
    @destructive
    def page_setup(
        self,
        sheet,
        orientation: str | None = None,
        paper: str | None = None,
        fit_to_pages_wide: int | None = None,
        fit_to_pages_tall: int | None = None,
        margins: dict[str, float] | None = None,
        print_area: str | None = None,
        center_horizontally: bool | None = None,
        center_vertically: bool | None = None,
        print_titles_rows: str | None = None,
        print_titles_cols: str | None = None,
        doc_id: str | None = None,
    ):
        ps = self._ws(sheet, doc_id).PageSetup
        if orientation is not None:
            ps.Orientation = _resolve_style(orientation, _ORIENTATION, "页面方向")
        if paper is not None:
            ps.PaperSize = _resolve_style(paper, _PAPER_SIZE, "纸张")
        if fit_to_pages_wide is not None or fit_to_pages_tall is not None:
            ps.Zoom = False  # FitToPages 与 Zoom 互斥：设 FitToPages 前必须关 Zoom
            if fit_to_pages_wide is not None:
                ps.FitToPagesWide = int(fit_to_pages_wide)
            if fit_to_pages_tall is not None:
                ps.FitToPagesTall = int(fit_to_pages_tall)
        if margins:
            # 边距单位磅（points），仅 left/right/top/bottom 四个方向可设
            for key in ("left", "right", "top", "bottom"):
                if key in margins:
                    setattr(ps, f"{key.capitalize()}Margin", margins[key])
        if print_area is not None:
            ps.PrintArea = print_area  # 空串清除打印区域
        if center_horizontally is not None:
            ps.CenterHorizontally = center_horizontally
        if center_vertically is not None:
            ps.CenterVertically = center_vertically
        if print_titles_rows is not None:
            ps.PrintTitleRows = print_titles_rows
        if print_titles_cols is not None:
            ps.PrintTitleColumns = print_titles_cols

    # --- 条件格式 ---
    @destructive
    def add_conditional_format(
        self,
        sheet,
        range_addr: str,
        rule: str,
        operator: str | None = None,
        value=None,
        value2=None,
        bg: str | None = None,
        fg: str | None = None,
        min_color: str | None = None,
        max_color: str | None = None,
        mid_color: str | None = None,
        doc_id: str | None = None,
    ):
        rule = rule.strip().lower()
        ws = self._ws(sheet, doc_id)
        rng = ws.Range(range_addr)
        if rule == "cell":
            if operator is None or value is None:
                raise InvalidArgumentError("cell 规则必须给 operator 和 value")
            op = _resolve_style(operator, _COND_OPERATOR, "条件格式运算符")
            if op in (1, 2) and value2 is None:  # between/not_between 需要 Formula2
                raise InvalidArgumentError("between/not_between 必须给 value2")
            fc = rng.FormatConditions.Add(1, op, value, value2)  # xlCellValue
            if bg is not None:
                fc.Interior.Color = _rgb(bg)
            if fg is not None:
                fc.Font.Color = _rgb(fg)
        elif rule == "databar":
            fc = rng.FormatConditions.Add(4)  # xlDatabar → 返回 Databar 对象
            if bg is not None:
                fc.BarColor.Color = _rgb(bg)  # Databar 实心填充色走 BarColor
        elif rule == "colorscale":
            if min_color is None or max_color is None:
                raise InvalidArgumentError("colorscale 必须给 min_color 和 max_color")
            n = 3 if mid_color else 2
            cs = rng.FormatConditions.AddColorScale(n)
            cs.ColorScaleCriteria(1).FormatColor.Color = _rgb(min_color)
            if mid_color:
                cs.ColorScaleCriteria(2).FormatColor.Color = _rgb(mid_color)
            cs.ColorScaleCriteria(n).FormatColor.Color = _rgb(max_color)
        else:
            raise InvalidArgumentError(
                f"未知条件格式规则: {rule!r}（可选: cell/databar/colorscale）"
            )

    # --- 基础三件套 ---
    @destructive
    def set_row_height(self, sheet, row, height: float, doc_id: str | None = None):
        self._ws(sheet, doc_id).Rows(row).RowHeight = height

    @destructive
    def set_number_format(self, sheet, range_addr: str, fmt: str, doc_id: str | None = None):
        self._ws(sheet, doc_id).Range(range_addr).NumberFormat = fmt

    @destructive
    def autofit(
        self,
        sheet,
        range_addr: str | None = None,
        columns: bool = True,
        rows: bool = True,
        doc_id: str | None = None,
    ):
        ws = self._ws(sheet, doc_id)
        target = ws.Range(range_addr) if range_addr else ws.UsedRange
        if columns:
            target.Columns.AutoFit()
        if rows:
            target.Rows.AutoFit()

    # --- 生命周期 ---
    def quit(self, force: bool = False):
        """退出 Excel 会话。

        own 句柄（本库启动的实例）直接退；连到既有 Office 实例默认拒绝
        （不夺走用户正用的窗口），确需退出传 force=True。实例已退（进程
        结束）视为已退出返回 True，不误报失败。
        """
        # 库改全局状态（DisplayAlerts），释放前还原原值
        if not self._owned and not force:
            with suppress(Exception):  # 仅兜底还原，失败不掩盖拒绝语义
                self.app.DisplayAlerts = self._saved_alerts
            raise ComOperationError("连接的是既有 Excel 实例，拒绝退出；确需退出请传 force=True")
        try:
            # P1-3：直接退自持句柄（不重连 ROT 里其它实例），避免误关别人的窗口
            self.app.DisplayAlerts = self._saved_alerts
            self.app.Quit()
        except Exception as e:  # noqa: BLE001 — com_error/断连异常统一走 liveness 判定
            if not core.doc_alive(self.app):
                return True  # 已退出：liveness 探针证实进程已结束
            raise ComOperationError(f"退出 Excel 失败: {e}") from e
