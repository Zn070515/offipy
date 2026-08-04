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
    call("word", "page_setup", orientation="landscape", paper="a4", top_margin=72, bottom_margin=72)
    ps = _word().ActiveDocument.PageSetup
    assert ps.Orientation == 1  # wdOrientLandscape
    assert ps.PaperSize == 7  # wdPaperA4
    assert ps.TopMargin == 72
