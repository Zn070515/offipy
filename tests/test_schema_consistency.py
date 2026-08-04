"""P1-2 operation schema 单一来源：schema ↔ server/CLI/MCP 三入口防漂移。

核心断言：
- schema 声明的 op 集合 ↔ server 白名单 _OPS 双向一致；destructive ↔
  _DESTRUCTIVE_OPS 一致
- schema 每个 op 的参数名 ↔ App 方法签名参数名一致（签名是类型权威）
- schema 每个 op 都有 MCP 工具（quit 除外）；每个 MCP 工具都映射回 schema op
- MCP 工具 readonly/destructive 注解与 schema 标志一致
"""

import inspect

from offipy import excel, mcp_server, ppt, schema, server, word

_APP_CLASSES = {"excel": excel.ExcelApp, "word": word.WordApp, "ppt": ppt.PptApp}


# --- schema ↔ server 白名单 ---


def test_schema_ops_match_server_whitelist():
    for app in schema.apps():
        assert server._OPS[app] == schema.ops(app)


def test_schema_destructive_matches_server():
    for app in schema.apps():
        assert server._DESTRUCTIVE_OPS[app] == schema.destructive_ops(app)


def test_server_ops_are_covered_by_schema():
    # 反向：server 白名单不会出现 schema 之外的 op（新增 RPC 只改 schema 一处）
    assert set(server._OPS) == set(schema.apps())
    for app, ops in server._OPS.items():
        assert ops == schema.ops(app)


def test_schema_apps_stable_order():
    assert schema.apps() == ("excel", "word", "ppt")


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


def test_schema_flags_internally_consistent():
    for app in schema.apps():
        for op in schema.ops(app):
            spec = schema.spec(app, op)
            assert not (spec.readonly and spec.destructive), f"{app}.{op} 不能同时只读且破坏性"
            if spec.readonly:
                assert op not in schema.destructive_ops(app), f"{app}.{op} 只读但被标破坏性"


def test_overwrite_ops_are_destructive():
    # P2-4 destructive 确认系统化：任何带 overwrite 参数（会覆盖目标文件）的
    # op 必须标 destructive——确保没有文件写入 op 逃过破坏性标记（覆盖保护
    # 由 paths.ensure_writable 统一施加，破坏性标记是它被调用的前提）。
    for app in schema.apps():
        for op, sp in schema.OPS[app].items():
            if "overwrite" in sp.params:
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
                assert ann.destructive_hint == spec.destructive
