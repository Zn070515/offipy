"""从 schema.OPS 生成 docs/api/ 下的 API 参考页面（P2-5 文档站），中英双语。

运行：`uv run python scripts/gen_api_ref.py`（纯标准库，无外部依赖）。
生成的 markdown 供 mkdocs 渲染：中文 `docs/api/{app}.md`，英文 `docs/api/{app}.en.md`。
schema.py 的描述保持中文单源（MCP/CLI 的工具描述不变）；英文由本文件的
`_EN_DESC` 翻译层提供。新增 op 时，`_EN_DESC` 必须补对应英文描述——
`tests/test_gen_api_ref_en.py` 兜底断言全量覆盖。
"""

from pathlib import Path

from offipy import schema

APP_NAMES = {"excel": "Excel", "word": "Word", "ppt": "PowerPoint"}
DOCS_API = Path(__file__).resolve().parent.parent / "docs" / "api"


def _type_name(t) -> str:
    return {
        str: "str",
        bool: "bool",
        int: "int",
        float: "float",
        list: "list",
        dict: "dict",
        object: "any",
        type(None): "null",
    }.get(t, "any")


def _param_list(params: dict) -> str:
    if not params:
        return "_无参数_"
    return "、".join(f"`{k}: {_type_name(v)}`" for k, v in params.items())


def _flags(spec: schema.OpSpec) -> list[str]:
    out = []
    if spec.readonly:
        out.append("只读")
    if spec.destructive:
        out.append("会改动文档/应用状态")
    if spec.deprecated:
        out.append("已弃用（响应带 warning）")
    return out


def _render_op(app: str, op: str, spec: schema.OpSpec) -> str:
    lines = [
        f"### `{op}`",
        "",
        spec.description or "",
        "",
        f"- **参数**: {_param_list(spec.params)}",
    ]
    lines.append(f"- **返回**: `{spec.returns}`")
    flags = _flags(spec)
    lines.append(f"- **标志**: {'，'.join(flags) if flags else '普通操作'}")
    return "\n".join(lines)


def _render_app(app: str) -> str:
    title = APP_NAMES[app]
    ops = schema.OPS[app]
    body = "\n\n---\n\n".join(_render_op(app, op, spec) for op, spec in ops.items())
    return f"> [English]({app}.en.md)\n\n# {title} API\n\n{body}\n"


def _render_index() -> str:
    rows = []
    for app in schema.apps():
        count = len(schema.OPS[app])
        read = len(schema.readonly_ops(app))
        destr = len(schema.destructive_ops(app))
        rows.append(f"| [{APP_NAMES[app]}]({app}.md) | {count} | {read} | {destr} |")
    table = "\n".join(rows)
    return (
        "> [English](index.en.md)\n\n"
        "# API 参考\n\n"
        "本参考由 `scripts/gen_api_ref.py` 从 `schema.py` 单一来源生成，"
        "覆盖 server / CLI / MCP 三入口的同一批操作。\n\n"
        "| 应用 | 操作数 | 只读 | 改动状态 |\n"
        "| --- | --- | --- | --- |\n"
        f"{table}\n"
        "\n每个操作：`doc_id` 缺省走当前活动文档（Excel `bookN` / Word `docN` / PPT `presN`）；"
        "`expected_target` 用于破坏性操作的绑定校验。\n\n"
        "> 静态几何质量门禁与基线回归不经过 `schema.py`（纯解析、无 Office/无 COM），"
        "另见 [PPTX 质量审计](audit.md) 与 [基线回归](audit-baseline.md)。\n"
    )


# ---------------------------------------------------------------------------
# English translation layer
#
# schema.py 的描述是中文单源（MCP/CLI 工具描述直接用它），英文版由这里翻译。
# 键是 `(app, op)`；任何 op 缺英文描述时生成会直接报错，tests 也有覆盖断言。
# ---------------------------------------------------------------------------

_EN_DESC: dict[tuple[str, str], str] = {
    # --- excel ---
    ("excel", "new_book"): "Create a new blank workbook, set it active, return doc_id.",
    ("excel", "open_book"): "Open an existing .xlsx/.xls file, set it active, return doc_id.",
    ("excel", "close_book"): (
        "Close the workbook (doc_id defaults to the active one). With save=True it saves "
        "first (a never-saved document auto-saves to the user data directory without the Save As "
        "dialog) and returns the save path; with save=False nothing is saved, no dialog, "
        "returns null."
    ),
    ("excel", "save"): (
        "Save the workbook (doc_id defaults to the active one) and return the absolute path. "
        "If path is given, save-as to that path; otherwise save back to the original path "
        "(a never-saved document auto-saves to the user data directory without the Save As dialog); "
        "overwrite=True allows overwriting an existing file."
    ),
    ("excel", "save_pdf"): (
        "Export the workbook (doc_id defaults to the active one) to PDF at the given path; "
        "overwrite=True allows overwriting an existing file."
    ),
    (
        "excel",
        "add_sheet",
    ): "Add a new worksheet to the workbook (doc_id defaults to the active one) and name it.",
    ("excel", "set_cell"): "Write a cell value; sheet takes a sheet name or index, cell like 'A1'.",
    ("excel", "get_cell"): "Read a cell value; sheet takes a sheet name or index, cell like 'A1'.",
    (
        "excel",
        "set_range",
    ): "Write a 2-D list of values into range_addr (e.g. 'A1:C3') in one call.",
    (
        "excel",
        "set_col_width",
    ): "Set column width; col takes a column number (1-based) or a column letter.",
    ("excel", "format_cell"): (
        "Format a cell. bold/italic take booleans; size is the font size; bg/fg take "
        "'#RRGGBB'; align takes an Excel horizontal-alignment constant."
    ),
    (
        "excel",
        "merge_cells",
    ): "Merge range_addr (e.g. 'A1:B2') into one cell, keeping the value in the top-left.",
    ("excel", "unmerge_cells"): "Unmerge range_addr.",
    ("excel", "set_border"): (
        "Set borders for range_addr. side is all/outside/inside or "
        "left/top/bottom/right/inside-h/inside-v; style is "
        "continuous/dash/dash-dot/dash-dot-dot/dot/double/none/slant-dash-dot; weight is "
        "hairline/thin/medium/thick; color takes '#RRGGBB'."
    ),
    ("excel", "freeze_panes"): (
        "Freeze the rows above row `rows` and the columns left of column `cols`; "
        "rows=0 and cols=0 unfreezes."
    ),
    ("excel", "page_setup"): (
        "Print settings. orientation is portrait/landscape; paper is letter/a3/a4; "
        "fit_to_pages_wide/tall take integers; margins takes a dict (in points); "
        "print_area takes 'A1:C10'; center_horizontally/center_vertically take booleans; "
        "print_titles_rows like '$1:$2'."
    ),
    ("excel", "add_conditional_format"): (
        "Add conditional formatting to range_addr. rule is cell/databar/colorscale; "
        "cell needs operator+value, colorscale needs min_color/max_color."
    ),
    ("excel", "set_row_height"): "Set the height of a row (in points).",
    (
        "excel",
        "set_number_format",
    ): "Set the number format for range_addr, e.g. '#,##0.00' / '0.0%' / 'yyyy-mm-dd'.",
    ("excel", "autofit"): (
        "Auto-fit column widths / row heights for range_addr; without range_addr it fits "
        "the used range. columns/rows are boolean toggles."
    ),
    (
        "excel",
        "read_range",
    ): "Read the values of worksheet range_addr (e.g. 'A1:C3'), returning a 2-D list (rows → columns).",
    (
        "excel",
        "activate",
    ): "Set the given doc_id as the active target; subsequent ops with a default doc_id act on it.",
    (
        "excel",
        "list_docs",
    ): "List the open-document table: {doc_id: {name, path, active}} (only registered handles).",
    ("excel", "get_target"): (
        "Identity of the active workbook (app/doc_id/name/path); null if none. "
        "Pass doc_id to query a specific workbook."
    ),
    ("excel", "quit"): (
        "Quit the Excel session (close the application window). Refuses by default when "
        "attached to an existing Office instance; force=True overrides."
    ),
    # --- word ---
    ("word", "new_doc"): "Create a new blank document, set it active, return doc_id.",
    ("word", "open_doc"): "Open an existing .docx/.doc file, set it active, return doc_id.",
    ("word", "close_doc"): (
        "Close the document (doc_id defaults to the active one). With save=True it saves "
        "first (a never-saved document auto-saves to the user data directory without the Save As "
        "dialog) and returns the save path; with save=False nothing is saved, no dialog, "
        "returns null."
    ),
    ("word", "save"): (
        "Save the document (doc_id defaults to the active one) and return the absolute path. "
        "If path is given, save-as to that path; otherwise save back to the original path "
        "(a never-saved document auto-saves to the user data directory without the Save As dialog); "
        "overwrite=True allows overwriting an existing file."
    ),
    ("word", "save_pdf"): (
        "Export the document (doc_id defaults to the active one) to PDF at the given path; "
        "overwrite=True allows overwriting an existing file."
    ),
    ("word", "write"): "Append text at the end of the document (no newline).",
    ("word", "write_line"): "Append a line of text at the end of the document (auto newline).",
    (
        "word",
        "add_heading",
    ): "Add a heading line at the end of the document and apply the Heading style (level 1-3).",
    (
        "word",
        "add_table",
    ): "Add a rows x cols table at the end of the document, returning the current table count.",
    (
        "word",
        "set_table_cell",
    ): "Set the (row, col) cell text of the table_idx-th table (rows/cols are 1-based).",
    ("word", "format_text"): (
        "Set the text format of the paragraph-th paragraph (1-based). bold/italic take "
        "booleans; size is the font size; name is the font name; color takes '#RRGGBB'; "
        "underline is none/single/words/double/dotted/wavy; highlight is "
        "none/yellow/green/pink/red/blue/bright_green/turquoise."
    ),
    ("word", "format_paragraph"): (
        "Set the paragraph format of the paragraph-th paragraph (1-based). alignment is "
        "left/center/right/justify; line_spacing is single/1.5/double/at_least/exactly/"
        "multiple; space_before/space_after/left_indent/first_line_indent are in points."
    ),
    ("word", "set_header_text"): "Set the header text of the section-th section.",
    ("word", "set_footer_text"): "Set the footer text of the section-th section.",
    ("word", "add_page_number"): (
        "Insert a page number in the footer. alignment is left/center/right; optional "
        "color '#RRGGBB' and size (font size); clears any existing footer text."
    ),
    ("word", "page_setup"): (
        "Page setup. orientation is portrait/landscape; paper is letter/legal/a3/a4/a5; "
        "left/right/top/bottom_margin and gutter are in points."
    ),
    ("word", "insert_toc"): (
        "Insert a table of contents at the start of the document (based on heading styles; "
        "levels controls the deepest heading level included)."
    ),
    (
        "word",
        "update_toc",
    ): "Update the table-of-contents fields in the document (refresh page numbers after adding/removing headings).",
    (
        "word",
        "add_list",
    ): "Append a list of lines at the end of the document; style is bullet or numbered.",
    (
        "word",
        "merge_table_cells",
    ): "Merge cells from (start_row,start_col) to (end_row,end_col) in the table_idx-th table.",
    ("word", "set_table_border"): (
        "Set borders of the table_idx-th table. style is none/single/dot/double; weight is "
        "0.25pt/0.5pt/0.75pt/1pt/1.5pt/2.25pt/3pt/4.5pt/6pt; color takes '#RRGGBB'; sides is "
        "all/outside/inside or left/top/bottom/right/inside-h/inside-v."
    ),
    (
        "word",
        "set_table_col_width",
    ): "Set the width of the col-th column of the table_idx-th table (in points).",
    ("word", "set_table_row_height"): (
        "Set the height of the row-th row of the table_idx-th table (in points). "
        "rule is auto/at_least/exactly."
    ),
    ("word", "autofit_table"): "Auto-fit the table_idx-th table. behavior is content/window/fixed.",
    ("word", "find_replace"): (
        "Find and replace throughout the document. replace_all replaces all occurrences, "
        "otherwise only the first; match_case/whole_word are optional."
    ),
    (
        "word",
        "insert_image",
    ): "Insert an image at the end of the document. width/height in points (omitted to keep the original size).",
    ("word", "insert_page_break"): "Insert a page break at the end of the document.",
    ("word", "read_doc_text"): "Read the full document text (read-only, does not modify state).",
    (
        "word",
        "activate",
    ): "Set the given doc_id as the active target; subsequent ops with a default doc_id act on it.",
    (
        "word",
        "list_docs",
    ): "List the open-document table: {doc_id: {name, path, active}} (only registered handles).",
    ("word", "get_target"): (
        "Identity of the active document (app/doc_id/name/path); null if none. "
        "Pass doc_id to query a specific document."
    ),
    ("word", "quit"): (
        "Quit the Word session (close the application window). Refuses by default when "
        "attached to an existing Office instance; force=True overrides."
    ),
    # --- ppt ---
    ("ppt", "new_pres"): "Create a new blank presentation, set it active, return doc_id.",
    ("ppt", "open_pres"): "Open an existing .pptx file, set it active, return doc_id.",
    ("ppt", "save"): (
        "Save the presentation (doc_id defaults to the active one) and return the absolute "
        "path. If path is given, save-as to that path (.pptx); otherwise save back to the "
        "original path (a never-saved document auto-saves to the user data directory without the "
        "Save As dialog); overwrite=True allows overwriting an existing file."
    ),
    ("ppt", "save_pdf"): (
        "Export the presentation (doc_id defaults to the active one) to PDF at the given "
        "path; overwrite=True allows overwriting an existing file."
    ),
    ("ppt", "export_slides"): (
        "Export each slide of the presentation (doc_id defaults to the active one) as PNG "
        "into out_dir (slide_01.png…), for visual inspection/iteration. Default 1920x1080. "
        "Returns the list of file paths."
    ),
    ("ppt", "add_slide"): (
        "Add a slide at the end. layout is {1: Title, 2: Title and Content, 5: Title Only, "
        "12: Blank}, default 2. Returns the current total slide count."
    ),
    (
        "ppt",
        "set_title",
    ): "Set the title text of the slide_idx-th slide; auto-adds a text box when no title placeholder exists. Returns the shape ID actually modified.",
    (
        "ppt",
        "set_body",
    ): "Set the body placeholder text of the slide_idx-th slide; lines is a list of strings, one per line. Auto-adds a text box when no body placeholder exists. Returns the shape ID actually modified.",
    (
        "ppt",
        "set_notes",
    ): "Write the speaker notes of the slide_idx-th slide. Returns the shape ID actually modified.",
    (
        "ppt",
        "add_textbox",
    ): "Add a free text box on the slide_idx-th slide (coordinates in points).",
    ("ppt", "add_picture"): "Insert an image on the slide_idx-th slide (coordinates in points).",
    ("ppt", "read_slide_texts"): (
        "Read every text-capable shape on the slide_idx-th slide (including text inside groups), "
        "returning SlideTextRecord entries (shape_id/name/text/coordinates/placeholder/group path). "
        "include_empty=True also returns text shapes with empty text; recursive=False skips groups."
    ),
    ("ppt", "read_slide_summary"): (
        "Read the title/body/notes summary of each slide of the presentation (doc_id defaults "
        "to the active one) (read-only), returning [{index, title, body, notes}]."
    ),
    (
        "ppt",
        "activate",
    ): "Set the given doc_id as the active target; subsequent ops with a default doc_id act on it.",
    (
        "ppt",
        "list_docs",
    ): "List the open-document table: {doc_id: {name, path, active}} (only registered handles).",
    ("ppt", "get_target"): (
        "Identity of the active presentation (app/doc_id/name/path); null if none. "
        "Pass doc_id to query a specific presentation."
    ),
    ("ppt", "quit"): (
        "Quit the PowerPoint session (close the application window). Refuses by default when "
        "attached to an existing Office instance; force=True overrides."
    ),
}


def _en_desc(app: str, op: str) -> str:
    try:
        return _EN_DESC[(app, op)]
    except KeyError:
        missing = [
            f"{a}.{o}" for a in schema.apps() for o in schema.ops(a) if (a, o) not in _EN_DESC
        ]
        raise ValueError(f"缺少英文描述: {missing}") from None


def _en_param_list(params: dict) -> str:
    if not params:
        return "_none_"
    return ", ".join(f"`{k}: {_type_name(v)}`" for k, v in params.items())


def _en_flags(spec: schema.OpSpec) -> list[str]:
    out = []
    if spec.readonly:
        out.append("read-only")
    if spec.destructive:
        out.append("mutates document/app state")
    if spec.deprecated:
        out.append("deprecated (response carries a warning)")
    return out


def _render_op_en(app: str, op: str, spec: schema.OpSpec) -> str:
    lines = [
        f"### `{op}`",
        "",
        _en_desc(app, op),
        "",
        f"- **Parameters**: {_en_param_list(spec.params)}",
    ]
    lines.append(f"- **Returns**: `{spec.returns}`")
    flags = _en_flags(spec)
    lines.append(f"- **Flags**: {', '.join(flags) if flags else 'normal operation'}")
    return "\n".join(lines)


def _render_app_en(app: str) -> str:
    title = APP_NAMES[app]
    ops = schema.OPS[app]
    body = "\n\n---\n\n".join(_render_op_en(app, op, spec) for op, spec in ops.items())
    return f"> [中文]({app}.md)\n\n# {title} API\n\n{body}\n"


def _render_index_en() -> str:
    rows = []
    for app in schema.apps():
        count = len(schema.OPS[app])
        read = len(schema.readonly_ops(app))
        destr = len(schema.destructive_ops(app))
        rows.append(f"| [{APP_NAMES[app]}]({app}.en.md) | {count} | {read} | {destr} |")
    table = "\n".join(rows)
    return (
        "> [中文](index.md)\n\n"
        "# API Reference\n\n"
        "This reference is generated from the single source of truth `schema.py` by "
        "`scripts/gen_api_ref.py` and covers the same set of operations across the three "
        "entry points (server / CLI / MCP).\n\n"
        "| App | Operations | Read-only | Mutating |\n"
        "| --- | --- | --- | --- |\n"
        f"{table}\n"
        "\nEvery operation: `doc_id` defaults to the current active document (Excel `bookN` "
        "/ Word `docN` / PPT `presN`); `expected_target` provides target binding for "
        "destructive operations.\n\n"
        "> Static geometry quality gates and baseline regression do not go through `schema.py` "
        "(pure parsing, no Office/COM); see [PPTX Quality Audit](audit.en.md) and "
        "[Baseline Regression](audit-baseline.en.md).\n"
    )


def _guard_en_coverage() -> None:
    missing = [f"{a}.{o}" for a in schema.apps() for o in schema.ops(a) if (a, o) not in _EN_DESC]
    if missing:
        raise ValueError(f"缺少英文描述: {missing}")


def main() -> None:
    # 覆盖守卫：任何 op 缺英文描述，生成直接失败（tests 另有兜底断言）
    _guard_en_coverage()
    DOCS_API.mkdir(parents=True, exist_ok=True)
    (DOCS_API / "index.md").write_text(_render_index(), encoding="utf-8")
    (DOCS_API / "index.en.md").write_text(_render_index_en(), encoding="utf-8")
    for app in schema.apps():
        (DOCS_API / f"{app}.md").write_text(_render_app(app), encoding="utf-8")
        (DOCS_API / f"{app}.en.md").write_text(_render_app_en(app), encoding="utf-8")
    print(f"generated {len(schema.apps()) + 1} pages (zh + en) under docs/api/")


if __name__ == "__main__":
    main()
