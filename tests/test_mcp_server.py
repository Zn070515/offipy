"""MCP server 测试：stdio 握手 + 工具清单（不需要 Office/COM）。

覆盖两个入口：python -m offipy.mcp_server（模块直启）和
python -m offipy.cli mcp（office mcp 的等价路径）。
"""

import sys

import anyio
import pytest
from mcp import ClientSession, StdioServerParameters, stdio_client


def _entry_args(module: str) -> list[str]:
    return ["-m", module, "mcp"] if module == "offipy.cli" else ["-m", module]


def _list_tools(module: str):
    async def main():
        params = StdioServerParameters(command=sys.executable, args=_entry_args(module))
        async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
            init = await session.initialize()
            assert init.server_info.name == "offipy"
            result = await session.list_tools()
            return {t.name: t for t in result.tools}

    return anyio.run(main)


@pytest.mark.parametrize("module", ["offipy.mcp_server", "offipy.cli"])
def test_initialize_and_tool_names(module):
    tools = _list_tools(module)
    for expected in (
        "ppt_new_presentation",
        "ppt_set_title",
        "word_write_line",
        "word_add_heading",
        "excel_set_cell",
        "excel_format_cell",
        "excel_merge_cells",
        "excel_unmerge_cells",
        "excel_set_border",
        "excel_add_conditional_format",
        "excel_freeze_panes",
        "excel_page_setup",
        "excel_set_row_height",
        "excel_set_number_format",
        "excel_autofit",
    ):
        assert expected in tools


@pytest.mark.parametrize("module", ["offipy.mcp_server", "offipy.cli"])
def test_tool_input_schema_requires_typed_args(module):
    tools = _list_tools(module)
    schema = tools["ppt_open_presentation"].input_schema
    assert schema["required"] == ["path"]
    assert schema["properties"]["path"]["type"] == "string"

    body = tools["ppt_set_body"].input_schema
    assert "slide_idx" in body["properties"]
    assert "lines" in body["properties"]

    cond = tools["excel_add_conditional_format"].input_schema
    assert cond["required"] == ["sheet", "range_addr", "rule"]
    for prop in ("operator", "value", "value2", "bg", "fg", "min_color", "max_color", "mid_color"):
        assert prop in cond["properties"]

    page = tools["excel_page_setup"].input_schema
    assert page["required"] == ["sheet"]
    for prop in (
        "orientation",
        "paper",
        "fit_to_pages_wide",
        "fit_to_pages_tall",
        "margins",
        "print_area",
        "center_horizontally",
        "center_vertically",
        "print_titles_rows",
        "print_titles_cols",
    ):
        assert prop in page["properties"]
