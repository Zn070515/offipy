"""Word 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActiveDocument 定位当前文档。
"""

import os

from . import core

WD_FORMAT_PDF = 17
# wdStyleHeading1..3 的 COM 常量值（-2/-3/-4），避免依赖中文/英文样式名
_HEADING_STYLES = {1: -2, 2: -3, 3: -4}


class WordApp:
    def __init__(self, visible: bool = True):
        self.app, _ = core.ensure_app("word", visible=visible)
        self._doc = None

    # --- 文档 ---
    def new_doc(self):
        self._doc = self.app.Documents.Add()
        return self._doc

    def open_doc(self, path: str):
        self._doc = self.app.Documents.Open(path)
        return self._doc

    def active_doc(self):
        if self._doc is not None:
            return self._doc
        doc = self.app.ActiveDocument
        if doc is None:
            doc = self.app.Documents.Add()
            self._doc = doc
        return doc

    def close_doc(self, save: bool = True):
        doc = self.active_doc()
        if doc is not None:
            doc.Close(SaveChanges=save)
        self._doc = None

    def save(self, path: str | None = None):
        doc = self.active_doc()
        if path:
            doc.SaveAs2(os.path.abspath(path))
        else:
            doc.Save()

    def save_pdf(self, path: str):
        self.active_doc().SaveAs2(os.path.abspath(path), FileFormat=WD_FORMAT_PDF)

    # --- 内容 ---
    def write(self, text: str):
        self.active_doc().Content.InsertAfter(text)

    def write_line(self, text: str):
        self.active_doc().Content.InsertAfter(text + "\r\n")

    def add_heading(self, text: str, level: int = 1):
        self.write_line(text)
        doc = self.active_doc()
        style = _HEADING_STYLES.get(level, -2)
        doc.Paragraphs(doc.Paragraphs.Count).Style = style

    def add_table(self, rows: int, cols: int):
        doc = self.active_doc()
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd：折叠到文末
        doc.Tables.Add(rng, rows, cols)
        return doc.Tables.Count

    def set_table_cell(self, table_idx: int, row: int, col: int, text: str):
        self.active_doc().Tables(table_idx).Cell(row, col).Range.Text = text

    def quit(self):
        core.quit_app("word")
