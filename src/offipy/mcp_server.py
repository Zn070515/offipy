"""MCP server：把 offipy 的 Office COM 会话操作暴露为标准 MCP 工具。

薄适配器：所有 COM 生命周期/会话管理仍由常驻的 8890 server 承担，
这里只做协议转换——每个 MCP 工具调用转成 client.request() 的 HTTP 调用。
Claude Desktop 等 MCP 客户端通过 stdio 拉起本进程（`office mcp` 或
`python -m offipy.mcp_server`），即可驱动真实 Word/Excel/PowerPoint：
窗口实时可见、状态跨调用保持，等同用户在 Office 里亲自操作当前文档。

注意：本进程绝不向 stdout 打印任何东西（stdio 传输协议占用 stdout）。
"""

from mcp.server import MCPServer

from . import __version__
from .client import request


def _call(app: str, op: str, **kwargs):
    """转成 8890 server 调用；失败抛 RuntimeError 让模型看到原因。"""
    try:
        resp = request(app, op, **kwargs)
    except SystemExit as e:
        raise RuntimeError(str(e) or f"offipy server 启动失败 (exit {e.code})") from None
    if not resp.get("ok"):
        raise RuntimeError(resp.get("error", "未知错误"))
    return resp.get("result")


server = MCPServer(
    name="offipy",
    title="offipy Office COM 自动化",
    description=(
        "会话式驱动真实 Microsoft Word/Excel/PowerPoint。所有操作作用在用户当前"
        "激活的文档/工作簿/演示文稿上（跨进程按 ActiveDocument/ActiveWorkbook/"
        "ActivePresentation 定位），窗口实时可见、状态跨调用保持。操作完的产物"
        "是原生 Office 文件，可继续在 Office 里手改。"
    ),
    version=__version__,
    log_level="WARNING",
)

# ---------------------------------------------------------------- PowerPoint


@server.tool(
    title="新建演示文稿",
    description="在 PowerPoint 中新建空白演示文稿，之后的操作都作用在它上面。",
)
def ppt_new_presentation() -> str:
    return str(_call("ppt", "new_pres"))


@server.tool(
    title="打开演示文稿",
    description="在 PowerPoint 中打开现有 .pptx，并把它设为当前文稿。",
)
def ppt_open_presentation(path: str) -> str:
    return str(_call("ppt", "open_pres", path=path))


@server.tool(
    title="保存演示文稿",
    description="保存当前演示文稿。给 path 则另存到该路径（.pptx）。",
)
def ppt_save(path: str | None = None) -> str:
    return str(_call("ppt", "save", path=path))


@server.tool(
    title="导出 PDF",
    description="把当前演示文稿导出为 PDF 到指定路径。",
)
def ppt_save_pdf(path: str) -> str:
    return str(_call("ppt", "save_pdf", path=path))


@server.tool(
    title="逐页导出 PNG",
    description=(
        "把当前演示文稿每一页导出为 PNG 到 out_dir（slide_01.png…），"
        "供视觉检查/迭代。默认 1920x1080。返回文件路径列表。"
    ),
)
def ppt_export_slides(out_dir: str, width: int = 1920, height: int = 1080) -> list[str]:
    return _call("ppt", "export_slides", out_dir=out_dir, width=width, height=height)


@server.tool(
    title="添加幻灯片",
    description=(
        "在末尾添加一张幻灯片。layout 取 {1:标题, 2:标题+文本, 5:仅标题, 12:空白}，"
        "默认 2。返回当前总页数。"
    ),
)
def ppt_add_slide(layout: int = 2) -> int:
    return _call("ppt", "add_slide", layout=layout)


@server.tool(
    title="设置标题",
    description="设置第 slide_idx 张幻灯片的标题文本。",
)
def ppt_set_title(slide_idx: int, text: str) -> str:
    return str(_call("ppt", "set_title", slide_idx=slide_idx, text=text))


@server.tool(
    title="设置正文",
    description="设置第 slide_idx 张幻灯片的正文占位符文本；lines 为逐行字符串列表。",
)
def ppt_set_body(slide_idx: int, lines: str | list[str]) -> str:
    return str(_call("ppt", "set_body", slide_idx=slide_idx, lines=lines))


@server.tool(
    title="设置演讲者备注",
    description="写入第 slide_idx 张幻灯片的演讲者备注。",
)
def ppt_set_notes(slide_idx: int, text: str) -> str:
    return str(_call("ppt", "set_notes", slide_idx=slide_idx, text=text))


@server.tool(
    title="添加文本框",
    description="在 slide_idx 页添加自由文本框（坐标单位为磅）。",
)
def ppt_add_textbox(
    slide_idx: int, left: float, top: float, width: float, height: float, text: str
) -> str:
    return str(
        _call(
            "ppt",
            "add_textbox",
            slide_idx=slide_idx,
            left=left,
            top=top,
            width=width,
            height=height,
            text=text,
        )
    )


@server.tool(
    title="添加图片",
    description="在 slide_idx 页插入图片（坐标单位为磅）。",
)
def ppt_add_picture(
    slide_idx: int, path: str, left: float, top: float, width: float, height: float
) -> str:
    return str(
        _call(
            "ppt",
            "add_picture",
            slide_idx=slide_idx,
            path=path,
            left=left,
            top=top,
            width=width,
            height=height,
        )
    )


# -------------------------------------------------------------------- Word


@server.tool(
    title="新建文档",
    description="在 Word 中新建空白文档，之后的操作都作用在它上面。",
)
def word_new_document() -> str:
    return str(_call("word", "new_doc"))


@server.tool(
    title="打开文档",
    description="在 Word 中打开现有 .docx/.doc，并把它设为当前文档。",
)
def word_open_document(path: str) -> str:
    return str(_call("word", "open_doc", path=path))


@server.tool(
    title="保存文档",
    description="保存当前文档。给 path 则另存到该路径。",
)
def word_save(path: str | None = None) -> str:
    return str(_call("word", "save", path=path))


@server.tool(
    title="导出 PDF",
    description="把当前 Word 文档导出为 PDF 到指定路径。",
)
def word_save_pdf(path: str) -> str:
    return str(_call("word", "save_pdf", path=path))


@server.tool(
    title="写入文本",
    description="在文档末尾追加文本（不换行）。",
)
def word_write(text: str) -> str:
    return str(_call("word", "write", text=text))


@server.tool(
    title="写入一行",
    description="在文档末尾追加一行文本（自动换行）。",
)
def word_write_line(text: str) -> str:
    return str(_call("word", "write_line", text=text))


@server.tool(
    title="添加标题",
    description="在文档末尾添加标题行并应用 Heading 样式（level 1-3）。",
)
def word_add_heading(text: str, level: int = 1) -> str:
    return str(_call("word", "add_heading", text=text, level=level))


@server.tool(
    title="添加表格",
    description="在文档末尾添加 rows x cols 表格，返回当前表格数。",
)
def word_add_table(rows: int, cols: int) -> int:
    return _call("word", "add_table", rows=rows, cols=cols)


@server.tool(
    title="设置表格单元格",
    description="设置第 table_idx 个表格的 (row, col) 单元格文本（行列 1 基）。",
)
def word_set_table_cell(table_idx: int, row: int, col: int, text: str) -> str:
    return str(_call("word", "set_table_cell", table_idx=table_idx, row=row, col=col, text=text))


@server.tool(
    title="关闭文档",
    description="关闭当前 Word 文档，save=True 则保存。",
)
def word_close_document(save: bool = True) -> str:
    return str(_call("word", "close_doc", save=save))


# ------------------------------------------------------------------- Excel


@server.tool(
    title="新建工作簿",
    description="在 Excel 中新建空白工作簿，之后的操作都作用在它上面。",
)
def excel_new_workbook() -> str:
    return str(_call("excel", "new_book"))


@server.tool(
    title="打开工作簿",
    description="在 Excel 中打开现有 .xlsx/.xls，并把它设为当前工作簿。",
)
def excel_open_workbook(path: str) -> str:
    return str(_call("excel", "open_book", path=path))


@server.tool(
    title="保存工作簿",
    description="保存当前工作簿。给 path 则另存到该路径。",
)
def excel_save(path: str | None = None) -> str:
    return str(_call("excel", "save", path=path))


@server.tool(
    title="导出 PDF",
    description="把当前工作簿导出为 PDF 到指定路径。",
)
def excel_save_pdf(path: str) -> str:
    return str(_call("excel", "save_pdf", path=path))


@server.tool(
    title="新建工作表",
    description="在活动工作簿中新建工作表并命名，返回该表对象。",
)
def excel_add_sheet(name: str) -> str:
    return str(_call("excel", "add_sheet", name=name))


@server.tool(
    title="设置单元格",
    description="写入单元格值；sheet 传表名或序号，cell 如 'A1'。",
)
def excel_set_cell(sheet: int | str, cell: str, value: str | int | float | bool) -> str:
    return str(_call("excel", "set_cell", sheet=sheet, cell=cell, value=value))


@server.tool(
    title="读取单元格",
    description="读取单元格的值；sheet 传表名或序号，cell 如 'A1'。",
)
def excel_get_cell(sheet: int | str, cell: str) -> str:
    return str(_call("excel", "get_cell", sheet=sheet, cell=cell))


@server.tool(
    title="批量写入区域",
    description="把二维值列表一次性写入 range_addr（如 'A1:C3'）。",
)
def excel_set_range(sheet: int | str, range_addr: str, values: list) -> str:
    return str(_call("excel", "set_range", sheet=sheet, range_addr=range_addr, values=values))


@server.tool(
    title="设置列宽",
    description="设置列宽；col 传列号（1 基）或列字母。",
)
def excel_set_col_width(sheet: int | str, col: int | str, width: float) -> str:
    return str(_call("excel", "set_col_width", sheet=sheet, col=col, width=width))


@server.tool(
    title="格式化单元格",
    description=(
        "格式化单元格。bold/italic 传布尔；size 字号；bg/fg 传 '#RRGGBB'；"
        "align 传 Excel 水平对齐常量（居中 -4108、左 -4131、右 -4152）。"
    ),
)
def excel_format_cell(
    sheet: int | str,
    cell: str,
    bold: bool | None = None,
    size: float | None = None,
    italic: bool | None = None,
    bg: str | None = None,
    fg: str | None = None,
    align: int | None = None,
) -> str:
    return str(
        _call(
            "excel",
            "format_cell",
            sheet=sheet,
            cell=cell,
            bold=bold,
            size=size,
            italic=italic,
            bg=bg,
            fg=fg,
            align=align,
        )
    )


@server.tool(
    title="关闭工作簿",
    description="关闭当前 Excel 工作簿，save=True 则保存。",
)
def excel_close_workbook(save: bool = True) -> str:
    return str(_call("excel", "close_book", save=save))


@server.tool(
    title="合并单元格",
    description="把 range_addr（如 'A1:B2'）合并为一个单元格，值保留在左上角。",
)
def excel_merge_cells(sheet: int | str, range_addr: str) -> str:
    return str(_call("excel", "merge_cells", sheet=sheet, range_addr=range_addr))


@server.tool(
    title="取消合并单元格",
    description="取消 range_addr 的合并。",
)
def excel_unmerge_cells(sheet: int | str, range_addr: str) -> str:
    return str(_call("excel", "unmerge_cells", sheet=sheet, range_addr=range_addr))


@server.tool(
    title="设置边框",
    description=(
        "给 range_addr 设置边框。side 取 all/outside/inside 或 "
        "left/top/bottom/right/inside-h/inside-v（逗号分隔）；style 取 "
        "continuous/dash/dash-dot/dash-dot-dot/dot/double/none/slant-dash-dot；"
        "weight 取 hairline/thin/medium/thick；color 传 '#RRGGBB'。"
    ),
)
def excel_set_border(
    sheet: int | str,
    range_addr: str,
    side: str = "all",
    style: str = "continuous",
    weight: str = "thin",
    color: str | None = None,
) -> str:
    return str(
        _call(
            "excel",
            "set_border",
            sheet=sheet,
            range_addr=range_addr,
            side=side,
            style=style,
            weight=weight,
            color=color,
        )
    )


@server.tool(
    title="设置条件格式",
    description=(
        "给 range_addr 加条件格式。rule 取 cell（单元格值规则，需 operator+value，可带 bg/fg）/ "
        "databar（数据条，可带 bg）/ colorscale（色阶，需 min_color/max_color，"
        "可带 mid_color 成三色）。"
        "operator 取 greater/less/between/equal/not_equal/greater_equal/less_equal/not_between。"
    ),
)
def excel_add_conditional_format(
    sheet: int | str,
    range_addr: str,
    rule: str,
    operator: str | None = None,
    value: str | int | float | None = None,
    value2: str | int | float | None = None,
    bg: str | None = None,
    fg: str | None = None,
    min_color: str | None = None,
    max_color: str | None = None,
    mid_color: str | None = None,
) -> str:
    return str(
        _call(
            "excel",
            "add_conditional_format",
            sheet=sheet,
            range_addr=range_addr,
            rule=rule,
            operator=operator,
            value=value,
            value2=value2,
            bg=bg,
            fg=fg,
            min_color=min_color,
            max_color=max_color,
            mid_color=mid_color,
        )
    )


@server.tool(
    title="冻结窗格",
    description="冻结 rows 行上方 + cols 列左侧；rows=0 且 cols=0 取消冻结。",
)
def excel_freeze_panes(sheet: int | str, rows: int = 0, cols: int = 0) -> str:
    return str(_call("excel", "freeze_panes", sheet=sheet, rows=rows, cols=cols))


@server.tool(
    title="页面设置",
    description=(
        "打印设置。orientation 取 portrait/landscape；paper 取 letter/a3/a4；"
        "fit_to_pages_wide/tall 传整数（设置后自动关 Zoom）；margins 传 "
        "{'left':..,'right':..,'top':..,'bottom':..}（单位磅）；print_area 传 'A1:C10'"
        "（空串清除）；center_horizontally/center_vertically 传布尔；"
        "print_titles_rows 如 '$1:$2'；print_titles_cols 如 '$A:$B'。"
    ),
)
def excel_page_setup(
    sheet: int | str,
    orientation: str | None = None,
    paper: str | None = None,
    fit_to_pages_wide: int | None = None,
    fit_to_pages_tall: int | None = None,
    margins: dict[str, float] | None = None,
    print_area: str | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
    print_titles_rows: str | None = None,
    print_titles_cols: str | None = None,
) -> str:
    return str(
        _call(
            "excel",
            "page_setup",
            sheet=sheet,
            orientation=orientation,
            paper=paper,
            fit_to_pages_wide=fit_to_pages_wide,
            fit_to_pages_tall=fit_to_pages_tall,
            margins=margins,
            print_area=print_area,
            center_horizontally=center_horizontally,
            center_vertically=center_vertically,
            print_titles_rows=print_titles_rows,
            print_titles_cols=print_titles_cols,
        )
    )


@server.tool(
    title="设置行高",
    description="设置某一行的高度（单位磅）。",
)
def excel_set_row_height(sheet: int | str, row: int, height: float) -> str:
    return str(_call("excel", "set_row_height", sheet=sheet, row=row, height=height))


@server.tool(
    title="设置数字格式",
    description="给 range_addr 设置数字格式，如 '#,##0.00' / '0.0%' / 'yyyy-mm-dd'。",
)
def excel_set_number_format(sheet: int | str, range_addr: str, fmt: str) -> str:
    return str(_call("excel", "set_number_format", sheet=sheet, range_addr=range_addr, fmt=fmt))


@server.tool(
    title="自动调整列宽行高",
    description=(
        "自动调整 range_addr 的列宽/行高；不传 range_addr 则调整已用区域（UsedRange）。"
        "columns/rows 为布尔开关，默认都调整，可只调列宽（rows=False）或只调行高（columns=False）。"
    ),
)
def excel_autofit(
    sheet: int | str,
    range_addr: str | None = None,
    columns: bool = True,
    rows: bool = True,
) -> str:
    return str(
        _call(
            "excel",
            "autofit",
            sheet=sheet,
            range_addr=range_addr,
            columns=columns,
            rows=rows,
        )
    )


def main():
    # stdio 传输：MCP 客户端（Claude Desktop 等）以子进程方式拉起并接管 stdin/stdout
    server.run()


if __name__ == "__main__":
    main()
