"""Excel 会话式自动化。

基于 core 的会话管理：每次调用重连同一个 Excel 实例，跨进程时通过
ActiveWorkbook 定位当前工作簿（即用户在 Excel 里当前激活的那个）。
"""

from . import core
from .paths import ensure_writable

# ExportAsFixedFormat 的类型常量
XL_TYPE_PDF = 0


def _parse_cell(cell: str):
    """把 'A1' 解析成 (row, col)，行列为 1 基。"""
    col_part = ""
    row_part = ""
    for ch in cell:
        if ch.isalpha():
            col_part += ch.upper()
        else:
            row_part += ch
    col = 0
    for ch in col_part:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_part), col


def _rgb(hex_color: str) -> int:
    """把 '#RRGGBB' 转成 Excel 的 BGR 整数颜色。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


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
        raise ValueError(f"无效边框侧: {side!r}")
    result = []
    for p in parts:
        if p not in _BORDER_INDEX:
            raise ValueError(f"未知边框侧: {p!r}（可选: {', '.join(_BORDER_INDEX)}）")
        result.append(_BORDER_INDEX[p])
    return result


def _resolve_style(name: str | None, table: dict[str, int], label: str) -> int:
    key = (name or "").strip().lower()
    if key not in table:
        raise ValueError(f"未知{label}: {name!r}（可选: {', '.join(table)}）")
    return table[key]


class ExcelApp:
    def __init__(self, visible: bool = True):
        self.app, self.created = core.ensure_app("excel", visible=visible)
        # 关闭所有提示（保存/覆盖/文件锁），避免模态对话框卡死单线程 server
        self._saved_alerts = self.app.DisplayAlerts  # 库改全局状态，释放时还原
        self.app.DisplayAlerts = False
        self._book = None

    # --- 工作簿 ---
    def new_book(self):
        self._book = self.app.Workbooks.Add()
        return self._book

    def open_book(self, path: str):
        self._book = self.app.Workbooks.Open(path)
        return self._book

    def active_book(self):
        # 会话语义（P1.2）：优先解析实时 ActiveWorkbook（用户当前激活的
        # 工作簿），仅当无活动工作簿时回退缓存句柄 + liveness probe。
        book = core.active_doc("excel", "ActiveWorkbook")
        if book is not None:
            self._book = book
            return book
        if self._book is not None and core.doc_alive(self._book):
            return self._book
        book = self.app.ActiveWorkbook
        if book is None:
            # 全新启动的 Excel 没有活动工作簿，自动新建一个保证可操作
            book = self.app.Workbooks.Add()
        self._book = book
        return book

    def close_book(self, save: bool = True):
        book = self.active_book()
        if book is not None:
            # Excel 的 xlDoNotSaveChanges=2（不是 False/0；0 不是合法值，会触发保存提示）
            book.Close(SaveChanges=-1 if save else 2)
        self._book = None

    def save(self, path: str | None = None, overwrite: bool = False):
        dest = ensure_writable(path, overwrite) if path else None
        book = self.active_book()
        if dest:
            # COM 的 SaveAs 不认正斜杠，必须规范为反斜杠绝对路径
            book.SaveAs(dest)
        else:
            book.Save()

    def save_pdf(self, path: str, overwrite: bool = False):
        dest = ensure_writable(path, overwrite)
        self.active_book().ExportAsFixedFormat(XL_TYPE_PDF, dest)

    # --- 工作表 ---
    def _ws(self, sheet):
        book = self.active_book()
        if isinstance(sheet, str):
            try:
                return book.Worksheets(sheet)
            except Exception:
                return book.Worksheets(int(sheet))
        return book.Worksheets(sheet)

    def add_sheet(self, name: str):
        ws = self.app.Worksheets.Add()
        ws.Name = name
        return ws

    # --- 单元格 ---
    def set_cell(self, sheet, cell: str, value):
        row, col = _parse_cell(cell)
        self._ws(sheet).Cells(row, col).Value = value

    def get_cell(self, sheet, cell: str):
        row, col = _parse_cell(cell)
        return self._ws(sheet).Cells(row, col).Value

    def set_range(self, sheet, range_addr: str, values):
        self._ws(sheet).Range(range_addr).Value = values

    def set_col_width(self, sheet, col, width):
        self._ws(sheet).Columns(col).ColumnWidth = width

    # --- 格式化 ---
    def format_cell(
        self, sheet, cell: str, bold=None, size=None, italic=None, bg=None, fg=None, align=None
    ):
        row, col = _parse_cell(cell)
        cell_obj = self._ws(sheet).Cells(row, col)
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
    def merge_cells(self, sheet, range_addr: str):
        self._ws(sheet).Range(range_addr).Merge()

    def unmerge_cells(self, sheet, range_addr: str):
        self._ws(sheet).Range(range_addr).UnMerge()

    # --- 边框 ---
    def set_border(
        self,
        sheet,
        range_addr: str,
        side: str = "all",
        style: str = "continuous",
        weight: str = "thin",
        color: str | None = None,
    ):
        ws = self._ws(sheet)
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
    def freeze_panes(self, sheet, rows: int = 0, cols: int = 0):
        if rows < 0 or cols < 0:
            raise ValueError(f"rows/cols 必须 ≥0，收到 rows={rows}, cols={cols}")
        ws = self._ws(sheet)
        ws.Activate()
        if rows == 0 and cols == 0:
            self.app.ActiveWindow.FreezePanes = False
        else:
            # 冻结 cell 左上方区域：选中 (rows+1, cols+1) 再开 FreezePanes
            ws.Cells(rows + 1, cols + 1).Select()
            self.app.ActiveWindow.FreezePanes = True

    # --- 打印设置 ---
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
    ):
        ps = self._ws(sheet).PageSetup
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
    ):
        rule = rule.strip().lower()
        ws = self._ws(sheet)
        rng = ws.Range(range_addr)
        if rule == "cell":
            if operator is None or value is None:
                raise ValueError("cell 规则必须给 operator 和 value")
            op = _resolve_style(operator, _COND_OPERATOR, "条件格式运算符")
            if op in (1, 2) and value2 is None:  # between/not_between 需要 Formula2
                raise ValueError("between/not_between 必须给 value2")
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
                raise ValueError("colorscale 必须给 min_color 和 max_color")
            n = 3 if mid_color else 2
            cs = rng.FormatConditions.AddColorScale(n)
            cs.ColorScaleCriteria(1).FormatColor.Color = _rgb(min_color)
            if mid_color:
                cs.ColorScaleCriteria(2).FormatColor.Color = _rgb(mid_color)
            cs.ColorScaleCriteria(n).FormatColor.Color = _rgb(max_color)
        else:
            raise ValueError(f"未知条件格式规则: {rule!r}（可选: cell/databar/colorscale）")

    # --- 基础三件套 ---
    def set_row_height(self, sheet, row, height: float):
        self._ws(sheet).Rows(row).RowHeight = height

    def set_number_format(self, sheet, range_addr: str, fmt: str):
        self._ws(sheet).Range(range_addr).NumberFormat = fmt

    def autofit(
        self, sheet, range_addr: str | None = None, columns: bool = True, rows: bool = True
    ):
        ws = self._ws(sheet)
        target = ws.Range(range_addr) if range_addr else ws.UsedRange
        if columns:
            target.Columns.AutoFit()
        if rows:
            target.Rows.AutoFit()

    # --- 生命周期 ---
    def quit(self):
        # 库改全局状态（DisplayAlerts），释放前还原原值
        self.app.DisplayAlerts = self._saved_alerts
        core.quit_app("excel")
