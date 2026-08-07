"""Word COM 集成测试：需要存活 Word（server 8890 持有），无则自动跳过。

有 Word 时，测试进程用 core.connect("word")（单实例应用 GetActiveObject）
直连读回断言；读不到就跑不起来 → pytestmark 兜底跳过。

P0-3 doc_id 权威：破坏性 op 需显式 doc_id，故每个测试捕获 new_doc 的
doc_id 并传给后续写入/格式 op。
"""

import os
import sys

import pytest

from offipy import core
from offipy.client import call
from offipy.exceptions import ComOperationError, InvalidArgumentError
from offipy.word import _rgb

pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(
        sys.platform != "win32" or not core.running("word"),
        reason="需要存活的 Word（server 8890 持有）",
    ),
]


def _word():
    return core.connect("word")


def test_format_text_bold_size_color():
    did = call("word", "new_doc")
    call("word", "write_line", text="重点内容", doc_id=did)
    call("word", "format_text", paragraph=1, bold=True, size=18, color="#2251FF", doc_id=did)
    font = _word().ActiveDocument.Paragraphs(1).Range.Font
    assert font.Bold in (True, -1)  # 早绑定下 VBA True 读回为 -1（int）
    assert font.Size == 18
    assert font.Color == _rgb("#2251FF")


def test_format_paragraph_alignment_line_spacing():
    did = call("word", "new_doc")
    call("word", "write_line", text="居中段", doc_id=did)
    call(
        "word",
        "format_paragraph",
        paragraph=1,
        alignment="center",
        line_spacing="double",
        doc_id=did,
    )
    fmt = _word().ActiveDocument.Paragraphs(1).Format
    assert fmt.Alignment == 1  # wdAlignParagraphCenter
    assert fmt.LineSpacingRule == 2  # wdLineSpaceDouble


def test_header_footer_text():
    did = call("word", "new_doc")
    call("word", "set_header_text", text="季度报告", doc_id=did)
    call("word", "set_footer_text", text="机密", doc_id=did)
    doc = _word().ActiveDocument
    assert doc.Sections(1).Headers(1).Range.Text.strip() == "季度报告"
    assert doc.Sections(1).Footers(1).Range.Text.strip() == "机密"


def test_add_page_number_center():
    did = call("word", "new_doc")
    call("word", "add_page_number", alignment="center", doc_id=did)
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    # S4：replace 改为直接插 PAGE 域 + 段落对齐（PageNumbers.Add 的对齐实探失效）
    assert hf.Range.ParagraphFormat.Alignment == 1  # wdAlignParagraphCenter
    assert any(f.Type == 33 for f in hf.Range.Fields)


def test_add_page_number_append_readback():
    did = call("word", "new_doc")
    call("word", "set_footer_text", text="公司名", doc_id=did)
    call("word", "add_page_number", mode="append", color="#2251FF", size=18, doc_id=did)
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    assert hf.Range.Text == "公司名1\r"  # 单段落符，无 \r 重复
    fld = next(f for f in hf.Range.Fields if f.Type == 33)
    assert fld.Result.Font.Color == _rgb("#2251FF")
    assert fld.Result.Font.Size == 18
    # 幂等：重复 append 不叠加
    call("word", "add_page_number", mode="append", doc_id=did)
    assert hf.Range.Text == "公司名1\r"
    assert len([f for f in hf.Range.Fields if f.Type == 33]) == 1


def test_add_page_number_standalone_readback():
    did = call("word", "new_doc")
    call("word", "set_footer_text", text="公司名", doc_id=did)
    call("word", "add_page_number", mode="standalone", color="#2251FF", size=18, doc_id=did)
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    assert hf.Range.Text == "公司名\t1\r"  # \t 是真实制表符，驱动 tab 分区
    fld = next(f for f in hf.Range.Fields if f.Type == 33)
    assert fld.Result.Font.Color == _rgb("#2251FF")
    assert fld.Result.Font.Size == 18
    # 幂等：重复 standalone 重建（先删既有 PAGE 域），不堆叠页码
    call("word", "add_page_number", mode="standalone", doc_id=did)
    assert hf.Range.Text == "公司名\t1\r"
    assert len([f for f in hf.Range.Fields if f.Type == 33]) == 1


def test_add_page_number_standalone_preserves_date_field():
    did = call("word", "new_doc")
    call("word", "set_footer_text", text="公司名", doc_id=did)
    # 先插一个 DATE 域（wdFieldDate=21）模拟既有非 PAGE 域
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    rng = hf.Range
    rng.MoveEnd(1, -1)  # wdCharacter：排除尾部段落符
    rng.Collapse(0)  # wdCollapseEnd
    rng.Fields.Add(rng, 21)  # wdFieldDate
    date_result = next(f for f in hf.Range.Fields if f.Type == 21).Result.Text
    # standalone 不得重写整段：DATE 域必须原样保留为活域
    call("word", "add_page_number", mode="standalone", doc_id=did)
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    types = [f.Type for f in hf.Range.Fields]
    assert 21 in types and 33 in types  # DATE 仍是活域，PAGE 域已加
    assert hf.Range.Text == f"公司名{date_result}\t1\r"


def test_add_page_number_standalone_alignment_switch_right_to_left():
    did = call("word", "new_doc")
    call("word", "set_footer_text", text="公司名", doc_id=did)
    call("word", "add_page_number", mode="standalone", alignment="right", doc_id=did)
    doc = _word().ActiveDocument
    hf = doc.Sections(1).Footers(1)
    text_w = round(
        doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin, 1
    )
    tabs = [(round(t.Position, 1), t.Alignment) for t in hf.Range.ParagraphFormat.TabStops]
    assert (text_w, 2) in tabs  # right 制表位存在
    assert hf.Range.Text == "公司名\t1\r"
    # 切到 left：旧右制表位被 ClearAll 清掉、遗留尾随制表符被移除，页码紧跟文本
    call("word", "add_page_number", mode="standalone", alignment="left", doc_id=did)
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    tabs = [(round(t.Position, 1), t.Alignment) for t in hf.Range.ParagraphFormat.TabStops]
    assert (text_w, 2) not in tabs  # 旧右制表位已清除
    assert "\t" not in hf.Range.Text  # 尾随制表符已移除（左模式无制表区）
    assert hf.Range.Text == "公司名1\r"


def test_add_page_number_standalone_left_follows_text():
    did = call("word", "new_doc")
    call("word", "set_footer_text", text="公司名", doc_id=did)
    call("word", "add_page_number", mode="standalone", alignment="left", doc_id=did)
    hf = _word().ActiveDocument.Sections(1).Footers(1)
    # 左模式无制表区：页码紧跟页脚文本自然左排（Word 丢弃 position 0 制表位）
    assert hf.Range.Text == "公司名1\r"
    assert len([f for f in hf.Range.Fields if f.Type == 33]) == 1


def test_page_setup_landscape_a4_margin():
    did = call("word", "new_doc")
    call(
        "word",
        "page_setup",
        orientation="landscape",
        paper="a4",
        top_margin=108,
        bottom_margin=144,
        doc_id=did,
    )
    ps = _word().ActiveDocument.PageSetup
    assert ps.Orientation == 1  # wdOrientLandscape
    assert ps.PaperSize == 7  # wdPaperA4
    assert ps.TopMargin == 108  # 非默认边距，确保赋值真的生效
    assert ps.BottomMargin == 144


def test_insert_and_update_toc():
    did = call("word", "new_doc")
    call("word", "add_heading", text="第一章", level=1, doc_id=did)
    doc = _word().ActiveDocument
    # add_heading 的标题样式须落在标题文本段（P1），而非 write_line 追加的空尾段（P2）；
    # 否则后续正文继承标题样式、目录（按标题样式收集）为空。
    assert doc.Paragraphs(1).Style.NameLocal in ("标题 1", "Heading 1")
    call("word", "write_line", text="正文内容", doc_id=did)
    call("word", "insert_toc", levels=3, doc_id=did)
    assert doc.TablesOfContents.Count == 1
    call("word", "update_toc", doc_id=did)  # 触发目录域刷新
    assert doc.TablesOfContents.Count == 1
    # 目录域实收标题（由标题样式驱动）；update_toc 后应能读到 "第一章"
    assert "第一章" in doc.TablesOfContents(1).Range.Text


def test_add_list_bullets():
    did = call("word", "new_doc")
    call("word", "add_list", lines=["甲", "乙", "丙"], style="bullet", doc_id=did)
    doc = _word().ActiveDocument
    # 空文档首段被复用来承接第一项；第 1-3 段都应是列表项。
    # 只查单段（ListType 只反映 range 首段），不能跨段取 range。
    # 列表项 ListType 为 2（bullet）/3（numbered），非列表段运行时返回 0。
    assert doc.Paragraphs(1).Range.ListFormat.ListType in (2, 3)
    assert doc.Paragraphs(2).Range.ListFormat.ListType in (2, 3)


def test_add_list_bullet_trailing_not_list():
    did = call("word", "new_doc")
    call("word", "add_list", lines=["甲", "乙", "丙"], style="bullet", doc_id=did)
    doc = _word().ActiveDocument
    # 列表项正确格式化（bullet=2 / numbered=3，实探值）
    for i in range(1, 4):
        assert doc.Paragraphs(i).Range.ListFormat.ListType in (2, 3)
    # 结构性空尾段（Count 恒含末段）保留但不再显示为列表项
    trailing = doc.Paragraphs(doc.Paragraphs.Count)
    assert trailing.Range.ListFormat.ListType == 0
    assert trailing.Range.Text == "\r"  # 段标记未被删除


def test_add_list_numbered_trailing_not_list():
    did = call("word", "new_doc")
    call("word", "add_list", lines=["一", "二"], style="numbered", doc_id=did)
    doc = _word().ActiveDocument
    for i in range(1, 3):
        assert doc.Paragraphs(i).Range.ListFormat.ListType in (2, 3)
    trailing = doc.Paragraphs(doc.Paragraphs.Count)
    assert trailing.Range.ListFormat.ListType == 0
    assert trailing.Range.Text == "\r"


def test_add_list_write_line_not_inherit_bullet():
    did = call("word", "new_doc")
    call("word", "add_list", lines=["甲"], style="bullet", doc_id=did)
    call("word", "write_line", text="后续正文", doc_id=did)
    doc = _word().ActiveDocument
    target = None
    for i in range(1, doc.Paragraphs.Count + 1):
        if doc.Paragraphs(i).Range.Text.strip() == "后续正文":
            target = doc.Paragraphs(i)
            break
    assert target is not None
    # 后续正文是普通段，不继承项目符号
    assert target.Range.ListFormat.ListType == 0


def test_add_list_repeated_no_accumulated_empty_bullets():
    did = call("word", "new_doc")
    call("word", "add_list", lines=["甲"], style="bullet", doc_id=did)
    call("word", "add_list", lines=["乙"], style="bullet", doc_id=did)
    doc = _word().ActiveDocument
    # 两批都是列表项
    assert doc.Paragraphs(1).Range.ListFormat.ListType in (2, 3)
    assert doc.Paragraphs(2).Range.ListFormat.ListType in (2, 3)
    # 恰好一个结构性空尾段，且不显示为列表项（不累积可见空项目符号）
    assert doc.Paragraphs.Count == 3
    trailing = doc.Paragraphs(doc.Paragraphs.Count)
    assert trailing.Range.ListFormat.ListType == 0
    assert trailing.Range.Text == "\r"


def test_merge_table_cells():
    did = call("word", "new_doc")
    call("word", "add_table", rows=2, cols=2, doc_id=did)
    call(
        "word",
        "merge_table_cells",
        table_idx=1,
        start_row=1,
        start_col=1,
        end_row=1,
        end_col=2,
        doc_id=did,
    )
    assert _word().ActiveDocument.Tables(1).Rows(1).Cells.Count == 1


def test_table_border_col_width_row_height_autofit():
    did = call("word", "new_doc")
    call("word", "add_table", rows=3, cols=2, doc_id=did)
    call(
        "word",
        "set_table_border",
        table_idx=1,
        style="single",
        color="#2251FF",
        sides="all",
        doc_id=did,
    )
    call("word", "set_table_col_width", table_idx=1, col=1, width=120, doc_id=did)
    call("word", "set_table_row_height", table_idx=1, row=1, height=30, doc_id=did)
    t = _word().ActiveDocument.Tables(1)
    assert t.Borders(1).LineStyle == 1  # wdLineStyleSingle
    assert t.Columns(1).Width == 120
    assert t.Rows(1).Height == 30
    call("word", "autofit_table", table_idx=1, behavior="content", doc_id=did)
    assert _word().ActiveDocument.Tables(1).Rows.Count == 3


def test_find_replace_all():
    did = call("word", "new_doc")
    call("word", "write_line", text="hello world hello", doc_id=did)
    call("word", "find_replace", find="hello", replace="hi", replace_all=True, doc_id=did)
    text = _word().ActiveDocument.Content.Text
    assert "hello" not in text
    assert text.count("hi") == 2


def test_insert_image_and_page_break(tmp_path):
    from PIL import Image

    img_path = tmp_path / "pixel.png"
    Image.new("RGB", (20, 20), "#2251FF").save(img_path)
    did = call("word", "new_doc")
    call("word", "insert_image", path=str(img_path), width=60, height=60, doc_id=did)
    call("word", "insert_page_break", doc_id=did)
    doc = _word().ActiveDocument
    assert doc.InlineShapes.Count == 1
    assert doc.InlineShapes(1).Width == 60
    assert "\x0c" in doc.Content.Text  # 分页符字符


# --- 边界用例（B6）：空输入 / 越界 / 目录缺失 / 受保护 / 兼容模式 / 路径 ---


def test_add_list_empty_rejected():
    did = call("word", "new_doc")
    with pytest.raises(InvalidArgumentError, match="lines 不能为空"):
        call("word", "add_list", lines=[], doc_id=did)
    with pytest.raises(InvalidArgumentError):
        call("word", "add_list", lines=[], style="numbered", doc_id=did)


def test_add_table_zero_dim_rejected():
    did = call("word", "new_doc")
    with pytest.raises(InvalidArgumentError, match="行列必须"):
        call("word", "add_table", rows=0, cols=2, doc_id=did)
    with pytest.raises(InvalidArgumentError):
        call("word", "add_table", rows=2, cols=0, doc_id=did)


def test_merge_cells_out_of_range_rejected():
    did = call("word", "new_doc")
    call("word", "add_table", rows=2, cols=2, doc_id=did)
    with pytest.raises(ComOperationError):
        call(
            "word",
            "merge_table_cells",
            table_idx=1,
            start_row=1,
            start_col=1,
            end_row=9,
            end_col=9,
            doc_id=did,
        )


def test_update_toc_without_toc_rejected():
    did = call("word", "new_doc")
    with pytest.raises(ComOperationError):
        call("word", "update_toc", doc_id=did)


def test_protected_document_rejects_write():
    did = call("word", "new_doc")
    call("word", "write_line", text="初始", doc_id=did)
    doc = _word().ActiveDocument
    doc.Protect(1, NoReset=True)  # wdAllowOnlyComments
    try:
        with pytest.raises(ComOperationError):
            call("word", "write_line", text="受保护写入", doc_id=did)
    finally:
        # 受保护文档上 Unprotect 会被「文档锁定」拒绝；直接丢弃文档清理
        call("word", "close_doc", doc_id=did, save=False)


def test_open_legacy_doc_compatibility_mode(tmp_path):
    w = _word()
    prev = w.DisplayAlerts
    w.DisplayAlerts = 0  # 抑制 .doc 兼容性检查等模态提示
    try:
        did = call("word", "new_doc")
        call("word", "write_line", text="旧格式内容", doc_id=did)
        p = tmp_path / "legacy.doc"
        _word().ActiveDocument.SaveAs(str(p), FileFormat=0)  # wdFormatDocument (97-2003)
        call("word", "close_doc", doc_id=did)
    finally:
        w.DisplayAlerts = prev
    did2 = call("word", "open_doc", path=str(p))
    # 兼容模式打开（CompatibilityMode < 15 即非 Word 2013+ 原生）
    assert _word().ActiveDocument.CompatibilityMode < 15
    assert "旧格式内容" in call("word", "read_doc_text", doc_id=did2)
    call("word", "close_doc", doc_id=did2)


def test_save_chinese_path(tmp_path):
    did = call("word", "new_doc")
    call("word", "write_line", text="中文路径", doc_id=did)
    out = call("word", "save", path=str(tmp_path / "中文报告.docx"), overwrite=True, doc_id=did)
    assert os.path.basename(out) == "中文报告.docx"
    assert os.path.exists(out)


def test_save_overlong_path_graceful(tmp_path):
    did = call("word", "new_doc")
    long_path = str(tmp_path / ("x" * 120) / ("长" * 120) / "报告.docx")  # 超 260 字符限制
    with pytest.raises((ComOperationError, OSError)):
        call("word", "save", path=long_path, overwrite=True, doc_id=did)
