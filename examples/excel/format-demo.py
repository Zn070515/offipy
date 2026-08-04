# examples/excel/format-demo.py
"""M4 Excel 格式能力示例：合并 / 边框 / 条件格式 / 冻结 / 打印 / 基础三件套。

运行：uv run python examples/excel/format-demo.py
（首次调用自动拉起 8890 server + Excel；产物保存到 out/excel-format-demo.xlsx）
"""

from offipy.client import call

call("excel", "new_book")

# 表头
call("excel", "set_cell", sheet=1, cell="A1", value="季度")
call("excel", "set_cell", sheet=1, cell="B1", value="营收(万元)")
call("excel", "set_cell", sheet=1, cell="C1", value="增长率")
call("excel", "set_cell", sheet=1, cell="D1", value="毛利")

# 数据
data = [
    ["Q1", 120, 0.08, 45],
    ["Q2", 150, 0.25, 58],
    ["Q3", 90, -0.4, 32],
    ["Q4", 210, 0.4, 88],
]
call("excel", "set_range", sheet=1, range_addr="A2:D5", values=data)

# 基础三件套
call("excel", "set_number_format", sheet=1, range_addr="B2:B5", fmt="#,##0")
call("excel", "set_number_format", sheet=1, range_addr="C2:C5", fmt="0.0%")
call("excel", "set_number_format", sheet=1, range_addr="D2:D5", fmt="#,##0")
call("excel", "set_row_height", sheet=1, row=1, height=24)
call("excel", "autofit", sheet=1, range_addr="A1:D5", rows=False)

# 边框 + 表头底色
call(
    "excel",
    "set_border",
    sheet=1,
    range_addr="A1:D5",
    side="all",
    style="continuous",
    weight="thin",
    color="#D0D7DE",
)
for cell in ("A1", "B1", "C1", "D1"):
    call("excel", "format_cell", sheet=1, cell=cell, bold=True, bg="#1F3A5F", fg="#FFFFFF")

# 条件格式：增长率正绿负红，营收数据条
call(
    "excel",
    "add_conditional_format",
    sheet=1,
    range_addr="C2:C5",
    rule="cell",
    operator="greater",
    value=0,
    bg="#C6EFCE",
    fg="#006100",
)
call(
    "excel",
    "add_conditional_format",
    sheet=1,
    range_addr="C2:C5",
    rule="cell",
    operator="less",
    value=0,
    bg="#FFC7CE",
    fg="#9C0006",
)
call("excel", "add_conditional_format", sheet=1, range_addr="B2:B5", rule="databar", bg="#2251FF")

# 冻结表头 + 打印设置
call("excel", "freeze_panes", sheet=1, rows=1, cols=0)
call(
    "excel",
    "page_setup",
    sheet=1,
    orientation="landscape",
    fit_to_pages_wide=1,
    center_horizontally=True,
)

call("excel", "save", path="out/excel-format-demo.xlsx")
print("已保存 out/excel-format-demo.xlsx")
