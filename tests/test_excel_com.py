"""Excel COM 集成测试：需要存活 Excel（server 8890 持有），无则自动跳过。

有 Excel 时，测试进程用 core.connect("excel")（单实例应用 GetActiveObject）
直连读回断言；读不到就跑不起来 → pytestmark 兜底跳过。
"""

import pytest

from offipy import core
from offipy.client import call

pytestmark = pytest.mark.skipif(
    not core.running("excel"),
    reason="需要存活的 Excel（server 8890 持有）",
)


def _excel():
    return core.connect("excel")


def test_merge_cells_keeps_topleft_value():
    call("excel", "new_book")
    call("excel", "set_cell", sheet=1, cell="A1", value="merged")
    call("excel", "merge_cells", sheet=1, range_addr="A1:B2")
    assert _excel().ActiveWorkbook.Sheets(1).Range("A1:B2").MergeCells is True
    assert call("excel", "get_cell", sheet=1, cell="A1") == "merged"
    call("excel", "unmerge_cells", sheet=1, range_addr="A1:B2")
    # 取消合并后 B2 可写可读
    call("excel", "set_cell", sheet=1, cell="B2", value="free")
    assert call("excel", "get_cell", sheet=1, cell="B2") == "free"


def test_set_border_thick_top():
    call("excel", "new_book")
    call(
        "excel",
        "set_border",
        sheet=1,
        range_addr="A1:C3",
        side="top",
        style="continuous",
        weight="thick",
        color="#FF0000",
    )
    ws = _excel().ActiveWorkbook.Sheets(1)
    b = ws.Range("A1:C3").Borders(8)  # xlEdgeTop
    assert b.LineStyle == 1  # xlContinuous
    assert b.Weight == 4  # xlThick
    assert b.Color == 255  # #FF0000 纯红 → BGR Long = 255


def test_freeze_panes_freeze_and_clear():
    call("excel", "new_book")
    call("excel", "freeze_panes", sheet=1, rows=1, cols=1)
    assert _excel().ActiveWindow.FreezePanes is True
    call("excel", "freeze_panes", sheet=1, rows=0, cols=0)
    assert _excel().ActiveWindow.FreezePanes is False


def test_page_setup_orientation_and_area():
    call("excel", "new_book")
    call(
        "excel",
        "page_setup",
        sheet=1,
        orientation="landscape",
        fit_to_pages_wide=1,
        center_horizontally=True,
        print_area="A1:C10",
    )
    ps = _excel().ActiveWorkbook.Sheets(1).PageSetup
    assert ps.Orientation == 2  # xlLandscape
    assert ps.FitToPagesWide == 1
    assert ps.CenterHorizontally is True
    assert ps.PrintArea == "$A$1:$C$10"


def test_conditional_format_cell_rule():
    call("excel", "new_book")
    call("excel", "set_cell", sheet=1, cell="A1", value=10)
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
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("A1:A5").FormatConditions(1)
    assert fc.Type == 1  # xlCellValue
    assert fc.Operator == 5  # xlGreater
    assert fc.Formula1 == "5"


def test_conditional_format_databar():
    call("excel", "new_book")
    call(
        "excel", "add_conditional_format", sheet=1, range_addr="B2:B5", rule="databar", bg="#2251FF"
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("B2:B5").FormatConditions(1)
    assert fc.Type == 4  # xlDatabar


def test_conditional_format_colorscale_three():
    call("excel", "new_book")
    call(
        "excel",
        "add_conditional_format",
        sheet=1,
        range_addr="C2:C5",
        rule="colorscale",
        min_color="#F8696B",
        max_color="#63BE7B",
        mid_color="#FFEB84",
    )
    fc = _excel().ActiveWorkbook.Sheets(1).Range("C2:C5").FormatConditions(1)
    assert fc.Type == 3  # xlColorScale
    assert fc.ColorScaleCriteria.Count == 3
