"""P1-2 operation schema 单一来源：schema ↔ server/CLI/MCP 三入口防漂移。

核心断言：
- schema 声明的 op 集合 ↔ server 白名单 _OPS 双向一致；destructive ↔
  _DESTRUCTIVE_OPS 一致
- schema 每个 op 的参数名 ↔ App 方法签名参数名一致（签名是类型权威）
- schema 每个 op 都有 MCP 工具（quit 除外）；每个 MCP 工具都映射回 schema op
- MCP 工具 readonly/destructive 注解与 schema 标志一致
"""

import inspect

from offipy import diagram, excel, mcp_server, ppt, schema, server, word

_APP_CLASSES = {
    "diagram": diagram.DiagramApp,
    "excel": excel.ExcelApp,
    "word": word.WordApp,
    "ppt": ppt.PptApp,
}


# --- schema ↔ server 白名单 ---


def test_schema_ops_match_server_whitelist():
    for app in schema.apps():
        assert server._OPS[app] == schema.ops(app)


def test_schema_requires_target_support_expected_target():
    # P0-3：导出 op（写文件系统）必须绑定目标——requires_target 自动继承
    # expected_target 支持（与 destructive 同一条路径）
    for app in schema.apps():
        for op in schema.ops(app):
            spec = schema.spec(app, op)
            want = bool(spec.destructive or spec.requires_target or spec.supports_expected_target)
            assert schema.supports_expected_target(app, op) == want, (
                f"{app}.{op} supports_expected_target 与标志不一致"
            )


def test_export_ops_require_target_binding():
    # P0-3：4 个导出 op 必须显式绑定目标（supports_expected_target True）；
    # 普通只读 op（read_range 类）不绑定
    export_ops = [
        ("excel", "save_pdf"),
        ("word", "save_pdf"),
        ("ppt", "save_pdf"),
        ("ppt", "export_slides"),
    ]
    for app, op in export_ops:
        assert schema.supports_expected_target(app, op), f"{app}.{op} 应绑定目标"
        assert schema.spec(app, op).requires_target
    assert schema.supports_expected_target("excel", "read_range") is False


def test_server_ops_are_covered_by_schema():
    # 反向：server 白名单不会出现 schema 之外的 op（新增 RPC 只改 schema 一处）
    assert set(server._OPS) == set(schema.apps())
    for app, ops in server._OPS.items():
        assert ops == schema.ops(app)


def test_schema_apps_stable_order():
    assert schema.apps() == ("excel", "word", "ppt", "diagram")


# --- schema 参数名 ↔ App 方法签名 ---


def test_schema_param_names_match_app_methods():
    for app in schema.apps():
        for op in schema.ops(app):
            spec = schema.spec(app, op)
            method = getattr(_APP_CLASSES[app], op)
            method_params = {
                p.name for p in inspect.signature(method).parameters.values() if p.name != "self"
            }
            assert set(spec.params) == method_params, (
                f"{app}.{op} schema 参数 {sorted(spec.params)} ≠ 方法参数 {sorted(method_params)}"
            )


def test_add_page_number_mode_in_schema():
    # S4 Task 2：schema 暴露 mode 参数（replace 默认 = legacy 行为），
    # 供 server/CLI/MCP 三入口与生成 API 使用；描述需说明默认值是 legacy 行为
    spec = schema.spec("word", "add_page_number")
    assert spec is not None
    assert spec.params["mode"] is str
    assert "replace" in spec.description and "legacy" in spec.description


def test_format_paragraph_line_spacing_union_in_schema():
    # S4 Task 3：line_spacing 以 tuple 编码 str|float 联合（CLI 保留字符串，
    # Python/JSON 可传数值）；描述需说明数值 1/1.5/2
    spec = schema.spec("word", "format_paragraph")
    assert spec is not None
    assert spec.params["line_spacing"] == (str, float)
    assert "数值 1/1.5/2" in spec.description


def test_schema_flags_internally_consistent():
    for app in schema.apps():
        for op in schema.ops(app):
            spec = schema.spec(app, op)
            assert not (spec.readonly and spec.destructive), f"{app}.{op} 不能同时只读且破坏性"
            assert not (spec.readonly and spec.requires_target), f"{app}.{op} 只读但需绑定目标"
            if spec.readonly:
                assert op not in schema.destructive_ops(app), f"{app}.{op} 只读但被标破坏性"


# 导出类 op：overwrite 保护的是「输出文件」（path 参数），不是 Office 文档——
# 它们不改文档，无需 doc_id 强制（P0-3：破坏性 = 改文档）。豁免
# 「overwrite→destructive」规则，但输出文件覆盖保护仍由 paths.ensure_writable 施加。
# diagram build 同属此类：只写输出 PPTX，无 Office 文档目标。
_OUTPUT_ONLY_OPS = {"save_pdf", "export_slides", "build"}


def test_overwrite_ops_are_destructive():
    # P2-4 destructive 确认系统化：任何带 overwrite 参数（会覆盖目标文件）的
    # op 必须标 destructive——确保没有文件写入 op 逃过破坏性标记（覆盖保护
    # 由 paths.ensure_writable 统一施加，破坏性标记是它被调用的前提）。
    for app in schema.apps():
        for op, sp in schema.OPS[app].items():
            if "overwrite" in sp.params and op not in _OUTPUT_ONLY_OPS:
                assert sp.destructive, f"{app}.{op} 带 overwrite 参数但未标 destructive"


# --- schema ↔ MCP 工具 ---


def _mcp_tools():
    return {t.name: t for t in mcp_server.server._tool_manager.list_tools()}


def test_every_schema_op_has_mcp_tool():
    tools = _mcp_tools()
    for app in schema.apps():
        for op in schema.ops(app):
            if op == "quit":
                continue  # quit 对整个会话太危险，MCP 不暴露
            name = mcp_server._tool_name(app, op)
            assert name in tools, f"schema 有 {app}.{op} 但缺 MCP 工具 {name}"
            assert hasattr(mcp_server, name)  # 模块级函数可直调（_invoke 契约）


def test_quit_not_exposed_to_mcp():
    tools = _mcp_tools()
    for app in schema.apps():
        assert f"{app}_quit" not in tools
        assert not hasattr(mcp_server, f"{app}_quit")


def test_every_mcp_tool_maps_back_to_schema_op():
    tools = _mcp_tools()
    for name in tools:
        app, _, _ = name.partition("_")
        assert app in schema.apps(), f"未知应用前缀: {name}"
        matched = [op for op in schema.ops(app) if mcp_server._tool_name(app, op) == name]
        assert len(matched) == 1, f"工具 {name} 无唯一对应 schema op（匹配到 {matched}）"


def test_mcp_tool_annotations_match_schema():
    tools = _mcp_tools()
    for app in schema.apps():
        for op in schema.ops(app):
            if op == "quit":
                continue
            spec = schema.spec(app, op)
            ann = tools[mcp_server._tool_name(app, op)].annotations
            assert ann is not None, f"{app}.{op} MCP 工具缺 annotations"
            if spec.readonly:
                assert ann.read_only_hint is True
                assert ann.destructive_hint is None
            else:
                assert ann.read_only_hint is None
                # P0-3：导出 op（requires_target）同样标记破坏性提示
                assert ann.destructive_hint == (spec.destructive or spec.requires_target)
