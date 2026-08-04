"""Word COM 集成测试：需要存活 Word（server 8890 持有），无则自动跳过。

有 Word 时，测试进程用 core.connect("word")（单实例应用 GetActiveObject）
直连读回断言；读不到就跑不起来 → pytestmark 兜底跳过。
"""
import pytest

from offipy import core
from offipy.client import call
from offipy.word import _rgb

pytestmark = pytest.mark.skipif(
    not core.running("word"),
    reason="需要存活的 Word（server 8890 持有）",
)


def _word():
    return core.connect("word")


def test_format_text_bold_size_color():
    call("word", "new_doc")
    call("word", "write_line", text="重点内容")
    call("word", "format_text", paragraph=1, bold=True, size=18, color="#2251FF")
    font = _word().ActiveDocument.Paragraphs(1).Range.Font
    assert font.Bold is True
    assert font.Size == 18
    assert font.Color == _rgb("#2251FF")


def test_format_paragraph_alignment_line_spacing():
    call("word", "new_doc")
    call("word", "write_line", text="居中段")
    call("word", "format_paragraph", paragraph=1, alignment="center", line_spacing="double")
    fmt = _word().ActiveDocument.Paragraphs(1).Format
    assert fmt.Alignment == 1  # wdAlignParagraphCenter
    assert fmt.LineSpacingRule == 2  # wdLineSpaceDouble


def test_header_footer_text():
    call("word", "new_doc")
    call("word", "set_header_text", text="季度报告")
    call("word", "set_footer_text", text="机密")
    doc = _word().ActiveDocument
    assert doc.Sections(1).Headers(1).Range.Text.strip() == "季度报告"
    assert doc.Sections(1).Footers(1).Range.Text.strip() == "机密"


def test_add_page_number_center():
    call("word", "new_doc")
    call("word", "add_page_number", alignment="center")
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    assert hf.PageNumbers.Count == 1
    assert hf.PageNumbers(1).Alignment == 1  # wdAlignPageNumberCenter


def test_page_setup_landscape_a4_margin():
    call("word", "new_doc")
    call(
        "word", "page_setup",
        orientation="landscape", paper="a4", top_margin=108, bottom_margin=144,
    )
    ps = _word().ActiveDocument.PageSetup
    assert ps.Orientation == 1  # wdOrientLandscape
    assert ps.PaperSize == 7  # wdPaperA4
    assert ps.TopMargin == 108  # 非默认边距，确保赋值真的生效
    assert ps.BottomMargin == 144


def test_insert_and_update_toc():
    call("word", "new_doc")
    call("word", "add_heading", text="第一章", level=1)
    call("word", "write_line", text="正文内容")
    call("word", "insert_toc", levels=3)
    doc = _word().ActiveDocument
    assert doc.TablesOfContents.Count == 1
    call("word", "update_toc")  # 不抛错即通过
    assert doc.TablesOfContents.Count == 1


def test_add_list_bullets():
    call("word", "new_doc")
    call("word", "add_list", lines=["甲", "乙", "丙"], style="bullet")
    doc = _word().ActiveDocument
    # 空文档首段（空段落）不被列表化；第 2-4 段是列表项。
    # 注意不能取跨段的 Range——ListFormat.ListType 只反映首段，跨到首段会回 -1。
    assert doc.Paragraphs(2).Range.ListFormat.ListType != -1  # -1 = wdListTypeNoList


def test_merge_table_cells():
    call("word", "new_doc")
    call("word", "add_table", rows=2, cols=2)
    call("word", "merge_table_cells", table_idx=1, start_row=1, start_col=1, end_row=1, end_col=2)
    assert _word().ActiveDocument.Tables(1).Rows(1).Cells.Count == 1


def test_table_border_col_width_row_height_autofit():
    call("word", "new_doc")
    call("word", "add_table", rows=3, cols=2)
    call("word", "set_table_border", table_idx=1, style="single", color="#2251FF", sides="all")
    call("word", "set_table_col_width", table_idx=1, col=1, width=120)
    call("word", "set_table_row_height", table_idx=1, row=1, height=30)
    t = _word().ActiveDocument.Tables(1)
    assert t.Borders(1).LineStyle == 1  # wdLineStyleSingle
    assert t.Columns(1).Width == 120
    assert t.Rows(1).Height == 30
    call("word", "autofit_table", table_idx=1, behavior="content")
    assert _word().ActiveDocument.Tables(1).Rows.Count == 3
