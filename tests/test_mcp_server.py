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
        "word_read_document_text",
        "ppt_read_slide_texts",
        "excel_read_range",
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


@pytest.mark.parametrize("module", ["offipy.mcp_server", "offipy.cli"])
def test_export_tools_expose_transport_params(module):
    # P0-3：导出 op（requires_target）暴露 expected_target/follow_active 传输参数
    tools = _list_tools(module)
    for name in ("ppt_save_pdf", "ppt_export_slides", "word_save_pdf", "excel_save_pdf"):
        schema = tools[name].input_schema
        assert "expected_target" in schema["properties"]
        assert "follow_active" in schema["properties"]


# --- _invoke 统一返回封装（mock _call，不需要 Office） ---


def test_server_description_mentions_target_binding():
    # 契约1：模型描述跟上 doc_id/follow_active/expected_target 设计，防焦点漂移
    from offipy import mcp_server

    desc = mcp_server.server.description
    assert "只读操作" in desc
    assert "doc_id" in desc
    assert "follow_active" in desc
    assert "expected_target" in desc


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

    def fake_call(app, op, request_id=None, **kw):
        captured[op] = kw
        return

    monkeypatch.setattr(mcp_server, "_call", fake_call)
    mcp_server.ppt_save(path="a.pptx", overwrite=True)
    assert captured["save"] == {"path": "a.pptx", "overwrite": True, "doc_id": None}
    mcp_server.word_save_pdf(path="b.pdf")
    assert captured["save_pdf"] == {"path": "b.pdf", "overwrite": False, "doc_id": None}
    mcp_server.excel_save()
    assert captured["save"] == {"path": None, "overwrite": False, "doc_id": None}


def test_tool_fn_passes_ctx_request_id(monkeypatch):
    # P1-4：MCP 框架注入 ctx，tool_fn 取其 request_id 透传 server（不出现在
    # tool schema）；无 ctx（测试直调）时 request_id=None。
    from offipy import mcp_server

    captured = {}

    def fake_call(app, op, request_id=None, **kw):
        captured["request_id"] = request_id
        return

    monkeypatch.setattr(mcp_server, "_call", fake_call)

    class _Ctx:
        request_id = "rid-ctx-7"

    mcp_server.ppt_add_slide(ctx=_Ctx())
    assert captured["request_id"] == "rid-ctx-7"

    mcp_server.ppt_add_slide()
    assert captured["request_id"] is None


# --- in-memory ClientSession（不起子进程，直连 MCPServer 协议层，P1-7） ---


def _run_inmemory(assert_fn):
    """在 in-memory transport 上跑 assert_fn(session)，验证 MCP 协议层行为。

    用 create_client_server_memory_streams 造双向内存流，server 端由
    _lowlevel_server.run 直接消费（与 run_stdio_async 同一路径），比
    stdio_client 拉起子进程更快、无需真实 stdio。
    """
    from mcp.shared.memory import create_client_server_memory_streams

    from offipy import mcp_server

    async def main():
        server = mcp_server.server
        async with (
            anyio.create_task_group() as tg,
            create_client_server_memory_streams() as (client_streams, server_streams),
        ):
            client_read, client_write = client_streams
            server_read, server_write = server_streams
            tg.start_soon(
                server._lowlevel_server.run,
                server_read,
                server_write,
                server._lowlevel_server.create_initialization_options(),
            )
            async with ClientSession(client_read, client_write) as session:
                await session.initialize()
                await assert_fn(session)

    anyio.run(main)


def test_inmemory_structured_return_passthrough(monkeypatch):
    from offipy import mcp_server

    # 真实 SlideTextRecord（0.10 契约）：list[SlideTextRecord] 而非 0.9 摘要 dict
    monkeypatch.setattr(
        mcp_server,
        "_call",
        lambda app, op, **kw: [
            {
                "shape_id": 2,
                "name": "Title 1",
                "text": "增长概览",
                "left": 36.0,
                "top": 21.6,
                "width": 648.0,
                "height": 90.0,
                "coordinate_space": "slide",
                "coordinate_unit": "pt",
                "is_placeholder": True,
                "placeholder_type": 1,
                "placeholder_type_name": "title",
                "parent_shape_id": None,
                "group_path": [],
            }
        ],
    )

    async def check(session):
        r = await session.call_tool("ppt_read_slide_texts", {"slide_idx": 1})
        assert r.is_error is False
        text = r.content[0].text
        assert '"shape_id": 2' in text and '"text": "增长概览"' in text
        assert '"placeholder_type_name": "title"' in text

    _run_inmemory(check)


def test_inmemory_read_slide_summary_passthrough(monkeypatch):
    from offipy import mcp_server

    # read_slide_summary 承接 0.9 摘要语义：返回 [{index, title, body, notes}]
    monkeypatch.setattr(
        mcp_server,
        "_call",
        lambda app, op, **kw: [{"index": 1, "title": "增长概览", "body": "MAU +18%", "notes": ""}],
    )

    async def check(session):
        r = await session.call_tool("ppt_read_slide_summary", {})
        assert r.is_error is False
        text = r.content[0].text
        assert '"index": 1' in text and '"title": "增长概览"' in text

    _run_inmemory(check)


def test_ppt_read_slide_texts_mcp_return_annotation():
    # P0：schema.returns="list[SlideTextRecord]" 必须映射到 list，
    # 否则 _RETURN_ANNOTATION.get 回退 object，MCP 结构化输出契约退化。
    import inspect

    from offipy import mcp_server

    annotation = inspect.signature(mcp_server.ppt_read_slide_texts).return_annotation
    assert annotation is list


def test_close_tools_return_annotation_is_str():
    # #33：schema.returns="str|null" 必须映射到 str，否则 _RETURN_ANNOTATION.get
    # 回退 object → close_* 的 MCP output_schema 退化为 None。运行时 _invoke 对
    # None 返回 "ok (op)"，返回值恒为 str，注解应结构化而非退化 object。
    import inspect

    from offipy import mcp_server

    for name in ("excel_close_book", "word_close_doc", "ppt_close_pres"):
        fn = getattr(mcp_server, name)
        assert inspect.signature(fn).return_annotation is str, name


def test_inmemory_int_and_void_and_args(monkeypatch):
    from offipy import mcp_server

    calls = {}

    def fake_call(app, op, request_id=None, **kw):
        calls[op] = kw
        if op == "add_slide":
            return 5
        if op in ("set_title", "set_body", "set_notes"):
            return 7  # B6：shape ID（int）
        return None

    monkeypatch.setattr(mcp_server, "_call", fake_call)

    async def check(session):
        r_int = await session.call_tool("ppt_add_slide")
        assert r_int.content[0].text == "5"
        r_void = await session.call_tool("ppt_new_presentation")
        assert r_void.content[0].text == "ok (new_pres)"
        r_args = await session.call_tool("ppt_set_title", {"slide_idx": 2, "text": "x"})
        assert r_args.content[0].text == "7"
        assert calls["set_title"] == {"slide_idx": 2, "text": "x", "doc_id": None}

    _run_inmemory(check)


def test_inmemory_tool_annotations_match_schema(monkeypatch):
    from offipy import mcp_server

    monkeypatch.setattr(mcp_server, "_call", lambda app, op, **kw: None)

    async def check(session):
        tools = {t.name: t for t in (await session.list_tools()).tools}
        ro = tools["ppt_read_slide_texts"].annotations
        assert ro is not None
        assert ro.read_only_hint is True
        assert ro.destructive_hint is None
        ds = tools["ppt_set_title"].annotations
        assert ds.read_only_hint is None
        assert ds.destructive_hint is True

    _run_inmemory(check)


def test_inmemory_call_error_maps_to_mcp_error(monkeypatch):
    from offipy import mcp_server

    def bad_call(app, op, **kw):
        raise RuntimeError("COM 失败: boom")

    monkeypatch.setattr(mcp_server, "_call", bad_call)

    async def check(session):
        r = await session.call_tool("excel_new_workbook")
        assert r.is_error is True

    _run_inmemory(check)
