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
        "word_format_text",
        "word_format_paragraph",
        "word_set_header_text",
        "word_set_footer_text",
        "word_add_page_number",
        "word_page_setup",
        "word_insert_toc",
        "word_update_toc",
        "word_add_list",
        "word_merge_table_cells",
        "word_set_table_border",
        "word_set_table_col_width",
        "word_set_table_row_height",
        "word_autofit_table",
        "word_find_replace",
        "word_insert_image",
        "word_insert_page_break",
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

    para = tools["word_format_paragraph"].input_schema
    assert para["required"] == ["paragraph"]
    for prop in (
        "alignment",
        "line_spacing",
        "space_before",
        "space_after",
        "left_indent",
        "first_line_indent",
    ):
        assert prop in para["properties"]

    fr = tools["word_find_replace"].input_schema
    assert fr["required"] == ["find", "replace"]
    assert fr["properties"]["replace_all"]["type"] == "boolean"

    tbl = tools["word_set_table_border"].input_schema
    assert tbl["required"] == ["table_idx"]
    for prop in ("style", "weight", "color", "sides"):
        assert prop in tbl["properties"]

    lst = tools["word_add_list"].input_schema
    assert lst["required"] == ["lines"]
    assert lst["properties"]["lines"]["type"] == "array"


@pytest.mark.parametrize("module", ["offipy.mcp_server", "offipy.cli"])
def test_save_tools_schema_exposes_overwrite(module):
    tools = _list_tools(module)
    for name in (
        "ppt_save",
        "ppt_save_pdf",
        "word_save",
        "word_save_pdf",
        "excel_save",
        "excel_save_pdf",
    ):
        schema = tools[name].input_schema
        assert schema["properties"]["overwrite"]["type"] == "boolean"
        assert "overwrite" not in schema.get("required", [])


# --- _invoke 统一返回封装（mock _call，不需要 Office） ---


def test_invoke_void_op_returns_ok_string(monkeypatch):
    from offipy import mcp_server

    monkeypatch.setattr(mcp_server, "_call", lambda app, op, **kw: None)
    assert mcp_server.ppt_new_presentation() == "ok (new_pres)"
    assert mcp_server.word_insert_page_break() == "ok (insert_page_break)"
    assert mcp_server.excel_new_workbook() == "ok (new_book)"


def test_invoke_value_op_structural_passthrough(monkeypatch):
    from offipy import mcp_server

    monkeypatch.setattr(mcp_server, "_call", lambda app, op, **kw: [1, 2])
    assert mcp_server.ppt_export_slides(out_dir="x") == [1, 2]

    monkeypatch.setattr(mcp_server, "_call", lambda app, op, **kw: 5)
    assert mcp_server.ppt_add_slide() == 5
    assert mcp_server.word_add_table(rows=2, cols=3) == 5


def test_save_tools_pass_overwrite_to_call(monkeypatch):
    from offipy import mcp_server

    captured = {}

    def fake_call(app, op, **kw):
        captured[op] = kw
        return None

    monkeypatch.setattr(mcp_server, "_call", fake_call)
    mcp_server.ppt_save(path="a.pptx", overwrite=True)
    assert captured["save"] == {"path": "a.pptx", "overwrite": True}
    mcp_server.word_save_pdf(path="b.pdf")
    assert captured["save_pdf"] == {"path": "b.pdf", "overwrite": False}
    mcp_server.excel_save()
    assert captured["save"] == {"path": None, "overwrite": False}
