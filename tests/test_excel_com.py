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
