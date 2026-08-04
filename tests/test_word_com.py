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
