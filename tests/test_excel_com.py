"""Excel COM 集成测试：需要存活 Excel（server 8890 持有），无则自动跳过。

有 Excel 时，测试进程用 core.connect("excel")（单实例应用 GetActiveObject）
直连读回断言；读不到就跑不起来 → pytestmark 兜底跳过。

P0-3 doc_id 权威：破坏性 op 需显式 doc_id，故每个测试捕获 new_book 的
doc_id 并传给后续写入/格式 op。
"""

import sys

import pytest

from offipy import core
from offipy.client import call
from offipy.excel import _rgb

pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(
        sys.platform != "win32" or not core.running("excel"),
        reason="需要存活的 Excel（server 8890 持有）",
    ),
]


def _excel():
    return core.connect("excel")


def test_merge_cells_keeps_topleft_value():
    did = call("excel", "new_book")
    call("excel", "set_cell", sheet=1, cell="A1", value="merged", doc_id=did)
    call("excel", "merge_cells", sheet=1, range_addr="A1:B2", doc_id=did)
    assert _excel().ActiveWorkbook.Sheets(1).Range("A1:B2").MergeCells is True
    assert call("excel", "get_cell", sheet=1, cell="A1", doc_id=did) == "merged"
    call("excel", "unmerge_cells", sheet=1, range_addr="A1:B2", doc_id=did)
    # 取消合并后 B2 可写可读
    call("excel", "set_cell", sheet=1, cell="B2", value="free", doc_id=did)
    assert call("excel", "get_cell", sheet=1, cell="B2", doc_id=did) == "free"


def test_set_border_thick_top():
    did = call("excel", "new_book")
    call(
        "excel",
        "set_border",
        sheet=1,
        range_addr="A1:C3",
        side="top",
        style="continuous",
        weight="thick",
        color="#FF0000",
        doc_id=did,
    )
    ws = _excel().ActiveWorkbook.Sheets(1)
    b = ws.Range("A1:C3").Borders(8)  # xlEdgeTop
    assert b.LineStyle == 1  # xlContinuous
    assert b.Weight == 4  # xlThick
    assert b.Color == 255  # #FF0000 纯红 → BGR Long = 255


def test_freeze_panes_freeze_and_clear():
    did = call("excel", "new_book")
    call("excel", "freeze_panes", sheet=1, rows=1, cols=1, doc_id=did)
    assert _excel().ActiveWindow.FreezePanes is True
    call("excel", "freeze_panes", sheet=1, rows=0, cols=0, doc_id=did)
    assert _excel().ActiveWindow.FreezePanes is False


def test_page_setup_orientation_and_area():
    did = call("excel", "new_book")
    call(
        "excel",
        "page_setup",
        sheet=1,
        orientation="landscape",
        fit_to_pages_wide=1,
        center_horizontally=True,
        print_area="A1:C10",
        doc_id=did,
    )
    ps = _excel().ActiveWorkbook.Sheets(1).PageSetup
    assert ps.Orientation == 2  # xlLandscape
    assert ps.FitToPagesWide == 1
    assert ps.CenterHorizontally is True
    assert ps.PrintArea == "$A$1:$C$10"


def test_conditional_format_cell_rule():
    did = call("excel", "new_book")
    call("excel", "set_cell", sheet=1, cell="A1", value=10, doc_id=did)
    call(
        "excel",
        "add_conditional_format",
        sheet=1,
        range_addr="A1:A5",
        rule="cell",
        operator="greater",
        value=5,
        bg="#FFC7CE",
        fg="#9C0006",
        doc_id=did,
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("A1:A5").FormatConditions(1)
    assert fc.Type == 1  # xlCellValue
    assert fc.Operator == 5  # xlGreater
    assert fc.Formula1 == "=5"  # Excel 回读 Formula1 会补 '=' 前缀


def test_conditional_format_databar():
    did = call("excel", "new_book")
    call(
        "excel",
        "add_conditional_format",
        sheet=1,
        range_addr="B2:B5",
        rule="databar",
        bg="#2251FF",
        doc_id=did,
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("B2:B5").FormatConditions(1)
    assert fc.Type == 4  # xlDatabar


def test_conditional_format_colorscale_three():
    did = call("excel", "new_book")
    call(
        "excel",
        "add_conditional_format",
        sheet=1,
        range_addr="C2:C5",
        rule="colorscale",
        min_color="#F8696B",
        max_color="#63BE7B",
        mid_color="#FFEB84",
        doc_id=did,
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("C2:C5").FormatConditions(1)
    assert fc.Type == 3  # xlColorScale
    assert fc.ColorScaleCriteria.Count == 3
    assert fc.ColorScaleCriteria(1).FormatColor.Color == _rgb("#F8696B")
    assert fc.ColorScaleCriteria(3).FormatColor.Color == _rgb("#63BE7B")


def test_conditional_format_colorscale_two():
    did = call("excel", "new_book")
    call(
        "excel",
        "add_conditional_format",
        sheet=1,
        range_addr="D2:D5",
        rule="colorscale",
        min_color="#F8696B",
        max_color="#63BE7B",
        doc_id=did,
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("D2:D5").FormatConditions(1)
    assert fc.Type == 3  # xlColorScale
    assert fc.ColorScaleCriteria.Count == 2


def test_set_row_height():
    did = call("excel", "new_book")
    call("excel", "set_row_height", sheet=1, row=1, height=30, doc_id=did)
    assert _excel().ActiveWorkbook.Sheets(1).Rows(1).RowHeight == 30


def test_set_number_format():
    did = call("excel", "new_book")
    call("excel", "set_cell", sheet=1, cell="A1", value=1234.5, doc_id=did)
    call("excel", "set_number_format", sheet=1, range_addr="A1:B1", fmt="#,##0.00", doc_id=did)
    assert _excel().ActiveWorkbook.Sheets(1).Range("A1").NumberFormat == "#,##0.00"


def test_autofit_widens_long_text_column():
    did = call("excel", "new_book")
    call(
        "excel",
        "set_cell",
        sheet=1,
        cell="A1",
        value="这是一段比较长的文本用于测试自动列宽",
        doc_id=did,
    )
    call("excel", "autofit", sheet=1, columns=True, rows=False, doc_id=did)
    assert _excel().ActiveWorkbook.Sheets(1).Columns(1).ColumnWidth > 10


def test_page_setup_margins():
    did = call("excel", "new_book")
    call("excel", "page_setup", sheet=1, margins={"left": 40, "top": 60}, doc_id=did)
    ps = _excel().ActiveWorkbook.Sheets(1).PageSetup
    assert ps.LeftMargin == 40
    assert ps.TopMargin == 60
