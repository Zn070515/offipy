"""M5 Word 能力示例：样式 / 页面结构 / 列表表格 / 文档辅助。

运行：uv run python examples/word/report-demo.py
（首次调用自动拉起 8890 server + Word；产物保存到 out/word-report-demo.docx）
"""

import os

from PIL import Image

from offipy.client import call

os.makedirs("out", exist_ok=True)

# 先造一张封面示意图（Pillow 是硬依赖）
img_path = "out/word-report-demo-cover.png"
Image.new("RGB", (800, 450), "#2251FF").save(img_path)

call("word", "new_doc")

# 封面：标题 + 副标题（样式系统）
call("word", "write_line", text="2026 半年度经营分析报告")
call("word", "format_text", paragraph=1, bold=True, size=26, color="#1F3A5F")
call("word", "format_paragraph", paragraph=1, alignment="center", space_after=12)
call("word", "write_line", text="编制：市场部 · 2026-08")
call("word", "format_paragraph", paragraph=2, alignment="center", space_after=24)
call("word", "insert_image", path=img_path, width=360, height=202)
# 空段分隔：让图片独占一段，避免后续标题与图片同段（Word 的 InsertAfter 追加到末段）
call("word", "write_line", text="")

# 章节标题 + 正文（标题样式供目录索引）
call("word", "add_heading", text="第一章 经营概况", level=1)
call("word", "write_line", text="本半年度整体营收保持增长，核心指标如下。")
# 实机段落计数：标题/副标题/图片/空段/章节标题/正文 → 正文为第 5 段（write_line 每次追加后留空尾段）
call(
    "word",
    "format_paragraph",
    paragraph=5,
    alignment="justify",
    line_spacing="1.5",
    first_line_indent=24,
)

# 列表
call(
    "word",
    "add_list",
    lines=["营收同比增长 18%", "毛利率提升 2.1 个百分点", "新客户签约 12 家"],
    style="bullet",
)

# 表格：合并表头 + 边框 + 列宽 + autofit
# 注意：列宽须在合并表头前设置——Word 整行合并后各列宽混合，
# 再访问单列 Width 会抛「混合的单元格宽度」COM 错误。
call("word", "add_table", rows=4, cols=3)
call("word", "set_table_col_width", table_idx=1, col=1, width=140)
call("word", "set_table_cell", table_idx=1, row=1, col=1, text="季度经营指标")
call("word", "merge_table_cells", table_idx=1, start_row=1, start_col=1, end_row=1, end_col=3)
for r, (a, b, c) in enumerate(
    [("营收(万元)", "1280", "1520"), ("毛利率", "32%", "34%"), ("净利率", "15%", "17%")],
    start=2,
):
    call("word", "set_table_cell", table_idx=1, row=r, col=1, text=a)
    call("word", "set_table_cell", table_idx=1, row=r, col=2, text=b)
    call("word", "set_table_cell", table_idx=1, row=r, col=3, text=c)
call("word", "set_table_border", table_idx=1, style="single", color="#9AA5B1", sides="all")
call("word", "set_table_row_height", table_idx=1, row=1, height=28)
call("word", "autofit_table", table_idx=1, behavior="window")

# 文档辅助 + 页面结构
call("word", "find_replace", find="季度", replace="半年度", replace_all=True)
call("word", "insert_page_break")
call("word", "set_header_text", text="2026 半年度经营分析报告 · 市场部")
call("word", "set_footer_text", text="机密文件")
call("word", "add_page_number", alignment="center")
call("word", "page_setup", orientation="landscape", paper="a4", top_margin=60, bottom_margin=60)

call("word", "save", path="out/word-report-demo.docx")
print("已保存 out/word-report-demo.docx")
