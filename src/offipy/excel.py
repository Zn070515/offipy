"""Excel 会话式自动化。

基于 core 的会话管理：每次调用重连同一个 Excel 实例，跨进程时通过
ActiveWorkbook 定位当前工作簿（即用户在 Excel 里当前激活的那个）。
"""

import os

from . import core

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


class ExcelApp:
    def __init__(self, visible: bool = True):
        self.app, self.created = core.ensure_app("excel", visible=visible)
        # 关闭所有提示（保存/覆盖/文件锁），避免模态对话框卡死单线程 server
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
        if self._book is not None:
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

    def save(self, path: str | None = None):
        book = self.active_book()
        if path:
            # COM 的 SaveAs 不认正斜杠，必须规范为反斜杠绝对路径
            book.SaveAs(os.path.abspath(path))
        else:
            book.Save()

    def save_pdf(self, path: str):
        self.active_book().ExportAsFixedFormat(XL_TYPE_PDF, os.path.abspath(path))

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

    # --- 生命周期 ---
    def quit(self):
        core.quit_app("excel")
