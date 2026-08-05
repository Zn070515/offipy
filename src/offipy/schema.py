"""operation schema：server / CLI / MCP 三入口的单一来源（P1-2）。

新增一个 RPC 只需两处：① 在 App 类实现方法（签名即参数类型与默认值），
② 在此登记一条 OpSpec（readonly/destructive/description/deprecated）。
server 白名单（_OPS/_DESTRUCTIVE_OPS）、CLI 参数校验/类型转换、MCP 工具
注册全部从此派生，不再三处手工同步。

参数签名以 App 方法为唯一权威（默认值/必填/类型都来自它）；schema 只声明
元数据与描述。一致性测试保证 schema 声明的 op 集合与参数名和 App 方法不漂移。

多文档（P2-2）：内容 op 统一带可选 doc_id（缺省走当前活动文档）；activate 切换
活动目标；list_docs 列出文档表。new_*/open_* 返回 doc_id（字符串）。
"""

from dataclasses import dataclass, field
from typing import Any

REQUIRED = object()  # 未用——参数必填性由 App 方法签名（无默认值）派生


@dataclass(frozen=True)
class OpSpec:
    """一个 RPC 操作的声明式元数据。"""

    description: str = ""
    readonly: bool = False  # 只读：不改任何文档/应用状态
    destructive: bool = False  # 会改动文档内容/状态（expected_target 绑定对象）
    supports_expected_target: bool = False  # 传输层额外支持 expected_target 绑定
    deprecated: bool = False  # P2-9 预留：已弃用 op，响应带 warning
    returns: str = "void"  # void/int/str/bool/list/dict/any（文档化用）
    params: dict[str, Any] = field(default_factory=dict)  # 参数类型（与 App 方法签名一致）


def apps() -> tuple[str, ...]:
    """已注册的应用名（excel/word/ppt），按定义顺序。"""
    return tuple(OPS)


def ops(app: str) -> frozenset[str]:
    """某应用的全部 op 名（server 白名单用）。"""
    return frozenset(OPS.get(app, {}))


def spec(app: str, op: str) -> OpSpec | None:
    """查 op 元数据；未知返回 None。"""
    return OPS.get(app, {}).get(op)


def destructive_ops(app: str) -> frozenset[str]:
    """某应用的所有破坏性 op（expected_target 绑定对象）。"""
    return frozenset(op for op, s in OPS.get(app, {}).items() if s.destructive)


def readonly_ops(app: str) -> frozenset[str]:
    """某应用的所有只读 op。"""
    return frozenset(op for op, s in OPS.get(app, {}).items() if s.readonly)


def supports_expected_target(app: str, op: str) -> bool:
    """该 op 是否暴露 expected_target 传输参数（P0-1）。

    destructive 自动继承（expected_target 用于防目标漂移）；个别非破坏性但
    可绑定目标的 op 可显式置 supports_expected_target=True。
    """
    s = spec(app, op)
    return bool(s and (s.destructive or s.supports_expected_target))


# =================================================================== Excel

OPS: dict[str, dict[str, OpSpec]] = {
    "excel": {
        "new_book": OpSpec(
            description="新建空白工作簿，设为活动，返回 doc_id。",
            returns="str",
        ),
        "open_book": OpSpec(
            description="打开现有 .xlsx/.xls，设为活动，返回 doc_id。",
            returns="str",
            params={"path": str},
        ),
        "close_book": OpSpec(
            description=(
                "关闭工作簿（doc_id 缺省为活动）。save=True 先保存（从未保存过则"
                "自动落盘用户数据目录，不弹另存为）并返回保存路径；save=False 不保存不弹窗，"
                "返回 null。"
            ),
            destructive=True,
            returns="str|null",
            params={"save": bool, "doc_id": str},
        ),
        "save": OpSpec(
            description=(
                "保存工作簿（doc_id 缺省为活动）并返回绝对路径。给 path 则另存到该"
                "路径；未给 path 则存回原路径（从未保存过自动落盘用户数据目录，不弹另存为）；"
                "overwrite=True 允许覆盖已存在文件。"
            ),
            destructive=True,
            returns="str",
            params={"path": str, "overwrite": bool, "doc_id": str},
        ),
        "save_pdf": OpSpec(
            description=(
                "把工作簿（doc_id 缺省为活动）导出为 PDF 到指定路径；"
                "overwrite=True 允许覆盖已存在文件。"
            ),
            params={"path": str, "overwrite": bool, "doc_id": str},
        ),
        "add_sheet": OpSpec(
            description="在工作簿（doc_id 缺省为活动）中新建工作表并命名。",
            destructive=True,
            params={"name": str, "doc_id": str},
        ),
        "set_cell": OpSpec(
            description="写入单元格值；sheet 传表名或序号，cell 如 'A1'。",
            destructive=True,
            params={"sheet": Any, "cell": str, "value": Any, "doc_id": str},
        ),
        "get_cell": OpSpec(
            description="读取单元格的值；sheet 传表名或序号，cell 如 'A1'。",
            readonly=True,
            returns="any",
            params={"sheet": Any, "cell": str, "doc_id": str},
        ),
        "set_range": OpSpec(
            description="把二维值列表一次性写入 range_addr（如 'A1:C3'）。",
            destructive=True,
            params={"sheet": Any, "range_addr": str, "values": Any, "doc_id": str},
        ),
        "set_col_width": OpSpec(
            description="设置列宽；col 传列号（1 基）或列字母。",
            destructive=True,
            params={"sheet": Any, "col": Any, "width": float, "doc_id": str},
        ),
        "format_cell": OpSpec(
            description=(
                "格式化单元格。bold/italic 传布尔；size 字号；"
                "bg/fg 传 '#RRGGBB'；align 传 Excel 水平对齐常量。"
            ),
            destructive=True,
            params={
                "sheet": Any,
                "cell": str,
                "bold": bool,
                "size": float,
                "italic": bool,
                "bg": str,
                "fg": str,
                "align": int,
                "doc_id": str,
            },
        ),
        "merge_cells": OpSpec(
            description="把 range_addr（如 'A1:B2'）合并为一个单元格，值保留在左上角。",
            destructive=True,
            params={"sheet": Any, "range_addr": str, "doc_id": str},
        ),
        "unmerge_cells": OpSpec(
            description="取消 range_addr 的合并。",
            destructive=True,
            params={"sheet": Any, "range_addr": str, "doc_id": str},
        ),
        "set_border": OpSpec(
            description=(
                "给 range_addr 设置边框。side 取 all/outside/inside 或 "
                "left/top/bottom/right/inside-h/inside-v；style 取 "
                "continuous/dash/dash-dot/dash-dot-dot/dot/double/none/slant-dash-dot；"
                "weight 取 hairline/thin/medium/thick；color 传 '#RRGGBB'。"
            ),
            destructive=True,
            params={
                "sheet": Any,
                "range_addr": str,
                "side": str,
                "style": str,
                "weight": str,
                "color": str,
                "doc_id": str,
            },
        ),
        "freeze_panes": OpSpec(
            description="冻结 rows 行上方 + cols 列左侧；rows=0 且 cols=0 取消冻结。",
            destructive=True,
            params={"sheet": Any, "rows": int, "cols": int, "doc_id": str},
        ),
        "page_setup": OpSpec(
            description=(
                "打印设置。orientation 取 portrait/landscape；paper 取 letter/a3/a4；"
                "fit_to_pages_wide/tall 传整数；margins 传字典（单位磅）；"
                "print_area 传 'A1:C10'；center_horizontally/center_vertically 传布尔；"
                "print_titles_rows 如 '$1:$2'。"
            ),
            destructive=True,
            params={
                "sheet": Any,
                "orientation": str,
                "paper": str,
                "fit_to_pages_wide": int,
                "fit_to_pages_tall": int,
                "margins": Any,
                "print_area": str,
                "center_horizontally": bool,
                "center_vertically": bool,
                "print_titles_rows": str,
                "print_titles_cols": str,
                "doc_id": str,
            },
        ),
        "add_conditional_format": OpSpec(
            description=(
                "给 range_addr 加条件格式。rule 取 cell/databar/colorscale；"
                "cell 需 operator+value，colorscale 需 min_color/max_color。"
            ),
            destructive=True,
            params={
                "sheet": Any,
                "range_addr": str,
                "rule": str,
                "operator": str,
                "value": Any,
                "value2": Any,
                "bg": str,
                "fg": str,
                "min_color": str,
                "max_color": str,
                "mid_color": str,
                "doc_id": str,
            },
        ),
        "set_row_height": OpSpec(
            description="设置某一行的高度（单位磅）。",
            destructive=True,
            params={"sheet": Any, "row": int, "height": float, "doc_id": str},
        ),
        "set_number_format": OpSpec(
            description="给 range_addr 设置数字格式，如 '#,##0.00' / '0.0%' / 'yyyy-mm-dd'。",
            destructive=True,
            params={"sheet": Any, "range_addr": str, "fmt": str, "doc_id": str},
        ),
        "autofit": OpSpec(
            description=(
                "自动调整 range_addr 的列宽/行高；不传 range_addr 则调整已用区域。"
                "columns/rows 为布尔开关。"
            ),
            destructive=True,
            params={"sheet": Any, "range_addr": str, "columns": bool, "rows": bool, "doc_id": str},
        ),
        "read_range": OpSpec(
            description="读取工作表 range_addr（如 'A1:C3'）的值，返回二维列表（行→列）。",
            readonly=True,
            returns="list",
            params={"sheet": Any, "range_addr": str, "doc_id": str},
        ),
        "activate": OpSpec(
            description="把指定 doc_id 设为活动目标，后续缺省 doc_id 的操作作用在它上面。",
            params={"doc_id": str},
        ),
        "list_docs": OpSpec(
            description="列出当前打开的文档表：{doc_id: {name, path, active}}（只报已登记句柄）。",
            readonly=True,
            returns="dict",
        ),
        "get_target": OpSpec(
            description=(
                "当前活动工作簿身份（app/doc_id/name/path）；无则返回 null。"
                "可传 doc_id 查询指定工作簿。"
            ),
            readonly=True,
            returns="dict",
            params={"doc_id": str},
        ),
        "quit": OpSpec(
            description=(
                "退出 Excel 会话（关闭应用窗口）。连接的是既有 Office 实例时默认拒绝"
                "（不夺走用户正用的窗口），force=True 强制退出。"
            ),
            destructive=True,
            params={"force": bool},
        ),
    },
    # ================================================================== Word
    "word": {
        "new_doc": OpSpec(
            description="新建空白文档，设为活动，返回 doc_id。",
            returns="str",
        ),
        "open_doc": OpSpec(
            description="打开现有 .docx/.doc，设为活动，返回 doc_id。",
            returns="str",
            params={"path": str},
        ),
        "close_doc": OpSpec(
            description=(
                "关闭文档（doc_id 缺省为活动）。save=True 先保存（从未保存过则"
                "自动落盘用户数据目录，不弹另存为）并返回保存路径；save=False 不保存不弹窗，"
                "返回 null。"
            ),
            destructive=True,
            returns="str|null",
            params={"save": bool, "doc_id": str},
        ),
        "save": OpSpec(
            description=(
                "保存文档（doc_id 缺省为活动）并返回绝对路径。给 path 则另存到该"
                "路径；未给 path 则存回原路径（从未保存过自动落盘用户数据目录，不弹另存为）；"
                "overwrite=True 允许覆盖已存在文件。"
            ),
            destructive=True,
            returns="str",
            params={"path": str, "overwrite": bool, "doc_id": str},
        ),
        "save_pdf": OpSpec(
            description=(
                "把文档（doc_id 缺省为活动）导出为 PDF 到指定路径；"
                "overwrite=True 允许覆盖已存在文件。"
            ),
            params={"path": str, "overwrite": bool, "doc_id": str},
        ),
        "write": OpSpec(
            description="在文档末尾追加文本（不换行）。",
            destructive=True,
            params={"text": str, "doc_id": str},
        ),
        "write_line": OpSpec(
            description="在文档末尾追加一行文本（自动换行）。",
            destructive=True,
            params={"text": str, "doc_id": str},
        ),
        "add_heading": OpSpec(
            description="在文档末尾添加标题行并应用 Heading 样式（level 1-3）。",
            destructive=True,
            params={"text": str, "level": int, "doc_id": str},
        ),
        "add_table": OpSpec(
            description="在文档末尾添加 rows x cols 表格，返回当前表格数。",
            destructive=True,
            returns="int",
            params={"rows": int, "cols": int, "doc_id": str},
        ),
        "set_table_cell": OpSpec(
            description="设置第 table_idx 个表格的 (row, col) 单元格文本（行列 1 基）。",
            destructive=True,
            params={"table_idx": int, "row": int, "col": int, "text": str, "doc_id": str},
        ),
        "format_text": OpSpec(
            description=(
                "设置第 paragraph 段（1 基）的文字格式。bold/italic 传布尔；size 字号；"
                "name 字体名；color 传 '#RRGGBB'；underline 取 "
                "none/single/words/double/dotted/wavy；highlight 取 "
                "none/yellow/green/pink/red/blue/bright_green/turquoise。"
            ),
            destructive=True,
            params={
                "paragraph": int,
                "bold": bool,
                "italic": bool,
                "size": float,
                "name": str,
                "color": str,
                "underline": str,
                "highlight": str,
                "doc_id": str,
            },
        ),
        "format_paragraph": OpSpec(
            description=(
                "设置第 paragraph 段（1 基）的段落格式。alignment 取 "
                "left/center/right/justify；line_spacing 取 "
                "single/1.5/double/at_least/exactly/multiple；"
                "space_before/space_after/left_indent/first_line_indent 单位磅。"
            ),
            destructive=True,
            params={
                "paragraph": int,
                "alignment": str,
                "line_spacing": str,
                "space_before": float,
                "space_after": float,
                "left_indent": float,
                "first_line_indent": float,
                "doc_id": str,
            },
        ),
        "set_header_text": OpSpec(
            description="设置第 section 节的页眉文本。",
            destructive=True,
            params={"text": str, "section": int, "doc_id": str},
        ),
        "set_footer_text": OpSpec(
            description="设置第 section 节的页脚文本。",
            destructive=True,
            params={"text": str, "section": int, "doc_id": str},
        ),
        "add_page_number": OpSpec(
            description=(
                "在页脚插入页码。alignment 取 left/center/right；"
                "可带 color '#RRGGBB' 和 size 字号（会清空既有页脚文本）。"
            ),
            destructive=True,
            params={"alignment": str, "color": str, "size": float, "doc_id": str},
        ),
        "page_setup": OpSpec(
            description=(
                "页面设置。orientation 取 portrait/landscape；"
                "paper 取 letter/legal/a3/a4/a5；"
                "left/right/top/bottom_margin 与 gutter 单位磅。"
            ),
            destructive=True,
            params={
                "orientation": str,
                "paper": str,
                "left_margin": float,
                "right_margin": float,
                "top_margin": float,
                "bottom_margin": float,
                "gutter": float,
                "doc_id": str,
            },
        ),
        "insert_toc": OpSpec(
            description="在文档开头插入目录（基于标题样式，levels 控制最深标题级别）。",
            destructive=True,
            params={"levels": int, "doc_id": str},
        ),
        "update_toc": OpSpec(
            description="更新文档中的目录域（新增/删除标题后刷新页码）。",
            destructive=True,
            params={"doc_id": str},
        ),
        "add_list": OpSpec(
            description=(
                "在文档末尾追加 lines 列表；style 取 bullet（项目符号）/ numbered（编号）。"
            ),
            destructive=True,
            params={"lines": list, "style": str, "doc_id": str},
        ),
        "merge_table_cells": OpSpec(
            description=(
                "把第 table_idx 个表格的 (start_row,start_col) 到 "
                "(end_row,end_col) 合并为一个单元格。"
            ),
            destructive=True,
            params={
                "table_idx": int,
                "start_row": int,
                "start_col": int,
                "end_row": int,
                "end_col": int,
                "doc_id": str,
            },
        ),
        "set_table_border": OpSpec(
            description=(
                "给第 table_idx 个表格设置边框。style 取 none/single/dot/double；"
                "weight 取 0.25pt/0.5pt/0.75pt/1pt/1.5pt/2.25pt/3pt/4.5pt/6pt；"
                "color 传 '#RRGGBB'；sides 取 all/outside/inside 或 "
                "left/top/bottom/right/inside-h/inside-v。"
            ),
            destructive=True,
            params={
                "table_idx": int,
                "style": str,
                "weight": str,
                "color": str,
                "sides": str,
                "doc_id": str,
            },
        ),
        "set_table_col_width": OpSpec(
            description="设置第 table_idx 个表格第 col 列的宽度（单位磅）。",
            destructive=True,
            params={"table_idx": int, "col": int, "width": float, "doc_id": str},
        ),
        "set_table_row_height": OpSpec(
            description=(
                "设置第 table_idx 个表格第 row 行的高度（单位磅）。rule 取 auto/at_least/exactly。"
            ),
            destructive=True,
            params={"table_idx": int, "row": int, "height": float, "rule": str, "doc_id": str},
        ),
        "autofit_table": OpSpec(
            description="自动调整第 table_idx 个表格。behavior 取 content/window/fixed。",
            destructive=True,
            params={"table_idx": int, "behavior": str, "doc_id": str},
        ),
        "find_replace": OpSpec(
            description=(
                "在全文执行查找替换。replace_all 为真则替换全部，否则只替换第一处；"
                "match_case/whole_word 可选。"
            ),
            destructive=True,
            params={
                "find": str,
                "replace": str,
                "match_case": bool,
                "whole_word": bool,
                "replace_all": bool,
                "doc_id": str,
            },
        ),
        "insert_image": OpSpec(
            description="在文档末尾插入图片。width/height 单位磅（省略则保持原尺寸）。",
            destructive=True,
            params={"path": str, "width": float, "height": float, "doc_id": str},
        ),
        "insert_page_break": OpSpec(
            description="在文档末尾插入分页符。",
            destructive=True,
            params={"doc_id": str},
        ),
        "read_doc_text": OpSpec(
            description="读取文档全文文本（只读，不修改状态）。",
            readonly=True,
            returns="str",
            params={"doc_id": str},
        ),
        "activate": OpSpec(
            description="把指定 doc_id 设为活动目标，后续缺省 doc_id 的操作作用在它上面。",
            params={"doc_id": str},
        ),
        "list_docs": OpSpec(
            description="列出当前打开的文档表：{doc_id: {name, path, active}}（只报已登记句柄）。",
            readonly=True,
            returns="dict",
        ),
        "get_target": OpSpec(
            description=(
                "当前活动文档身份（app/doc_id/name/path）；无则返回 null。"
                "可传 doc_id 查询指定文档。"
            ),
            readonly=True,
            returns="dict",
            params={"doc_id": str},
        ),
        "quit": OpSpec(
            description=(
                "退出 Word 会话（关闭应用窗口）。连接的是既有 Office 实例时默认拒绝"
                "（不夺走用户正用的窗口），force=True 强制退出。"
            ),
            destructive=True,
            params={"force": bool},
        ),
    },
    # ==================================================================== PPT
    "ppt": {
        "new_pres": OpSpec(
            description="新建空白演示文稿，设为活动，返回 doc_id。",
            returns="str",
        ),
        "open_pres": OpSpec(
            description="打开现有 .pptx，设为活动，返回 doc_id。",
            returns="str",
            params={"path": str},
        ),
        "save": OpSpec(
            description=(
                "保存演示文稿（doc_id 缺省为活动）并返回绝对路径。给 path 则另存到"
                "该路径（.pptx）；未给 path 则存回原路径（从未保存过自动落盘用户数据目录，"
                "不弹另存为）；overwrite=True 允许覆盖已存在文件。"
            ),
            destructive=True,
            returns="str",
            params={"path": str, "overwrite": bool, "doc_id": str},
        ),
        "save_pdf": OpSpec(
            description=(
                "把演示文稿（doc_id 缺省为活动）导出为 PDF 到指定路径；"
                "overwrite=True 允许覆盖已存在文件。"
            ),
            params={"path": str, "overwrite": bool, "doc_id": str},
        ),
        "export_slides": OpSpec(
            description=(
                "把演示文稿（doc_id 缺省为活动）每一页导出为 PNG 到 out_dir"
                "（slide_01.png…），供视觉检查/迭代。默认 1920x1080。返回文件路径列表。"
            ),
            returns="list",
            params={"out_dir": str, "width": int, "height": int, "doc_id": str},
        ),
        "add_slide": OpSpec(
            description=(
                "在末尾添加一张幻灯片。layout 取 "
                "{1:标题, 2:标题+文本, 5:仅标题, 12:空白}，默认 2。返回当前总页数。"
            ),
            destructive=True,
            returns="int",
            params={"layout": int, "doc_id": str},
        ),
        "set_title": OpSpec(
            description="设置第 slide_idx 张幻灯片的标题文本。",
            destructive=True,
            params={"slide_idx": int, "text": str, "doc_id": str},
        ),
        "set_body": OpSpec(
            description="设置第 slide_idx 张幻灯片的正文占位符文本；lines 为逐行字符串列表。",
            destructive=True,
            params={"slide_idx": int, "lines": Any, "doc_id": str},
        ),
        "set_notes": OpSpec(
            description="写入第 slide_idx 张幻灯片的演讲者备注。",
            destructive=True,
            params={"slide_idx": int, "text": str, "doc_id": str},
        ),
        "add_textbox": OpSpec(
            description="在 slide_idx 页添加自由文本框（坐标单位为磅）。",
            destructive=True,
            params={
                "slide_idx": int,
                "left": float,
                "top": float,
                "width": float,
                "height": float,
                "text": str,
                "doc_id": str,
            },
        ),
        "add_picture": OpSpec(
            description="在 slide_idx 页插入图片（坐标单位为磅）。",
            destructive=True,
            params={
                "slide_idx": int,
                "path": str,
                "left": float,
                "top": float,
                "width": float,
                "height": float,
                "doc_id": str,
            },
        ),
        "read_slide_texts": OpSpec(
            description=(
                "逐页读取演示文稿（doc_id 缺省为活动）的标题/正文/备注文本（只读），"
                "返回 [{index, title, body, notes}]。"
            ),
            readonly=True,
            returns="list",
            params={"doc_id": str},
        ),
        "activate": OpSpec(
            description="把指定 doc_id 设为活动目标，后续缺省 doc_id 的操作作用在它上面。",
            params={"doc_id": str},
        ),
        "list_docs": OpSpec(
            description="列出当前打开的文档表：{doc_id: {name, path, active}}（只报已登记句柄）。",
            readonly=True,
            returns="dict",
        ),
        "get_target": OpSpec(
            description=(
                "当前活动演示文稿身份（app/doc_id/name/path）；无则返回 null。"
                "可传 doc_id 查询指定演示文稿。"
            ),
            readonly=True,
            returns="dict",
            params={"doc_id": str},
        ),
        "quit": OpSpec(
            description=(
                "退出 PowerPoint 会话（关闭应用窗口）。连接的是既有 Office 实例时默认拒绝"
                "（不夺走用户正用的窗口），force=True 强制退出。"
            ),
            destructive=True,
            params={"force": bool},
        ),
    },
}
