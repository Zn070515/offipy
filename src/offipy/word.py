"""Word 会话式自动化。

基于 core 的会话管理；跨进程时通过 ActiveDocument 定位当前文档。
"""

import os

from . import core
from .paths import ensure_writable

WD_FORMAT_PDF = 17
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


def _rgb(hex_color: str) -> int:
    """把 '#RRGGBB' 转成 COM Long 颜色（Word 与 Excel 同用 COLORREF 公式）。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


def _resolve_style(name: str | None, table: dict[str, int], label: str) -> int:
    key = (name or "").strip().lower()
    if key not in table:
        raise ValueError(f"未知{label}: {name!r}（可选: {', '.join(table)}）")
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
        raise ValueError(f"无效表格边: {sides!r}")
    result = []
    for p in parts:
        if p not in _TABLE_SIDES:
            raise ValueError(f"未知表格边: {p!r}（可选: {', '.join(_TABLE_SIDES)}）")
        result.append(_TABLE_SIDES[p])
    return result


def _end_range(doc):
    """返回折叠到文档末尾的 Range（用于文末插入）。"""
    rng = doc.Content
    rng.Collapse(0)  # wdCollapseEnd
    return rng


class WordApp:
    def __init__(self, visible: bool = True):
        self.app, _ = core.ensure_app("word", visible=visible)
        self.app.DisplayAlerts = 0  # wdAlertsNone：抑制保存/覆盖等模态提示
        self._doc = None

    # --- 文档 ---
    def new_doc(self):
        self._doc = self.app.Documents.Add()
        return self._doc

    def open_doc(self, path: str):
        self._doc = self.app.Documents.Open(path)
        return self._doc

    def active_doc(self):
        # 会话语义（P1.2）：优先解析实时 ActiveDocument（用户当前激活的
        # 文档），仅当无活动文档时回退缓存句柄 + liveness probe。
        doc = core.active_doc("word", "ActiveDocument")
        if doc is not None:
            self._doc = doc
            return doc
        if self._doc is not None and core.doc_alive(self._doc):
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

    def save(self, path: str | None = None, overwrite: bool = False):
        dest = ensure_writable(path, overwrite) if path else None
        doc = self.active_doc()
        if dest:
            doc.SaveAs2(dest)
        else:
            doc.Save()

    def save_pdf(self, path: str, overwrite: bool = False):
        dest = ensure_writable(path, overwrite)
        self.active_doc().SaveAs2(dest, FileFormat=WD_FORMAT_PDF)

    # --- 内容 ---
    def write(self, text: str):
        self.active_doc().Content.InsertAfter(text)

    def write_line(self, text: str):
        self.active_doc().Content.InsertAfter(text + "\r\n")

    def add_heading(self, text: str, level: int = 1):
        self.write_line(text)
        doc = self.active_doc()
        style = _HEADING_STYLES.get(level, -2)
        # write_line = Content.InsertAfter(text + "\r\n")：文本落在末尾空段之前的 Count-1 段
        # （Count 是空尾段，实机验证）。若给 Count 上样式，随后正文会继承标题样式，
        # 目录（按标题样式收集）为空。
        doc.Paragraphs(doc.Paragraphs.Count - 1).Style = style

    def add_table(self, rows: int, cols: int):
        doc = self.active_doc()
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd：折叠到文末
        doc.Tables.Add(rng, rows, cols)
        return doc.Tables.Count

    def set_table_cell(self, table_idx: int, row: int, col: int, text: str):
        self.active_doc().Tables(table_idx).Cell(row, col).Range.Text = text

    # --- 样式系统：文字格式 ---
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
    ):
        font = self.active_doc().Paragraphs(paragraph).Range.Font
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
    def format_paragraph(
        self,
        paragraph: int,
        alignment: str | None = None,
        line_spacing: str | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        left_indent: float | None = None,
        first_line_indent: float | None = None,
    ):
        fmt = self.active_doc().Paragraphs(paragraph).Format
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
    def set_header_text(self, text: str, section: int = 1):
        self.active_doc().Sections(section).Headers(1).Range.Text = text

    def set_footer_text(self, text: str, section: int = 1):
        self.active_doc().Sections(section).Footers(1).Range.Text = text

    def add_page_number(
        self,
        alignment: str = "right",
        color: str | None = None,
        size: float | None = None,
    ):
        hf = self.active_doc().Sections(1).Footers(1)
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

    def page_setup(
        self,
        orientation: str | None = None,
        paper: str | None = None,
        left_margin: float | None = None,
        right_margin: float | None = None,
        top_margin: float | None = None,
        bottom_margin: float | None = None,
        gutter: float | None = None,
    ):
        ps = self.active_doc().PageSetup
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
    def insert_toc(self, levels: int = 3):
        doc = self.active_doc()
        doc.TablesOfContents.Add(
            doc.Range(0, 0), UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=levels
        )
        return doc.TablesOfContents.Count

    def update_toc(self):
        doc = self.active_doc()
        doc.TablesOfContents(1).Update()
        return doc.TablesOfContents.Count

    # --- 列表 ---
    def add_list(self, lines: list[str], style: str = "bullet"):
        doc = self.active_doc()
        start = doc.Paragraphs.Count + 1  # 第一个新段落的序号
        for line in lines:
            self.write_line(line)
        end = doc.Paragraphs.Count
        # Range 起点前移一个字符，落在上一段段末标记上，
        # 避免 ApplyBulletDefault 跳过 range 首段（Word 段落边界行为）
        rng = doc.Range(doc.Paragraphs(start).Range.Start - 1, doc.Paragraphs(end).Range.End)
        if style == "numbered":
            rng.ListFormat.ApplyNumberDefault()
        else:
            rng.ListFormat.ApplyBulletDefault()
        return len(lines)

    # --- 表格增强 ---
    def merge_table_cells(
        self, table_idx: int, start_row: int, start_col: int, end_row: int, end_col: int
    ):
        t = self.active_doc().Tables(table_idx)
        t.Cell(start_row, start_col).Merge(t.Cell(end_row, end_col))

    def set_table_border(
        self,
        table_idx: int,
        style: str = "single",
        weight: str | None = None,
        color: str | None = None,
        sides: str | None = None,
    ):
        t = self.active_doc().Tables(table_idx)
        const = _resolve_style(style, _LINE_STYLE, "线型")
        for idx in _resolve_table_sides(sides):
            b = t.Borders(idx)
            b.LineStyle = const
            if weight is not None:
                b.LineWidth = _resolve_style(weight, _TABLE_LINE_WIDTH, "线宽")
            if color is not None:
                b.Color = _rgb(color)

    def set_table_col_width(self, table_idx: int, col: int, width: float):
        self.active_doc().Tables(table_idx).Columns(col).Width = width

    def set_table_row_height(self, table_idx: int, row: int, height: float, rule: str = "at_least"):
        r = self.active_doc().Tables(table_idx).Rows(row)
        r.Height = height
        r.HeightRule = _resolve_style(rule, _ROW_HEIGHT_RULE, "行高规则")

    def autofit_table(self, table_idx: int, behavior: str = "content"):
        self.active_doc().Tables(table_idx).AutoFitBehavior(
            _resolve_style(behavior, _AUTOFIT, "自动调整")
        )

    # --- 文档辅助 ---
    def find_replace(
        self,
        find: str,
        replace: str,
        match_case: bool = False,
        whole_word: bool = False,
        replace_all: bool = True,
    ):
        f = self.active_doc().Content.Find
        f.Execute(
            FindText=find,
            ReplaceWith=replace,
            Replace=_REPLACE["all" if replace_all else "one"],
            Forward=True,
            MatchCase=match_case,
            MatchWholeWord=whole_word,
        )

    def insert_image(self, path: str, width: float | None = None, height: float | None = None):
        doc = self.active_doc()
        shape = doc.InlineShapes.AddPicture(os.path.abspath(path), Range=_end_range(doc))
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height
        return doc.InlineShapes.Count

    def insert_page_break(self):
        _end_range(self.active_doc()).InsertBreak(7)  # wdPageBreak

    def quit(self):
        core.quit_app("word")
