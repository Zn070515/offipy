"""MCP server：把 offipy 的 Office COM 会话操作暴露为标准 MCP 工具。

薄适配器：所有 COM 生命周期/会话管理仍由常驻的 8890 server 承担，
这里只做协议转换——每个 MCP 工具调用转成 client.request() 的 HTTP 调用。
Claude Desktop 等 MCP 客户端通过 stdio 拉起本进程（`offipy mcp` 或
`python -m offipy.mcp_server`），即可驱动真实 Word/Excel/PowerPoint：
窗口实时可见、状态跨调用保持，等同用户在 Office 里亲自操作当前文档。

工具集由 operation schema（P1-2）驱动注册：schema 声明每个 op 的存在、
描述与 readonly/destructive 元数据；参数签名（必填/默认值/类型）以 App
方法为唯一权威派生。新增 RPC 只改 schema.py + App 方法，MCP 工具自动
跟随。quit 不暴露（对整个会话太危险）；get_target 随 schema 透出。

注意：本进程绝不向 stdout 打印任何东西（stdio 传输协议占用 stdout）。
"""

import inspect
from typing import Any, cast

from mcp.server import MCPServer
from mcp.server.mcpserver import Context
from mcp_types import ToolAnnotations

from . import __version__, schema
from .client import request
from .excel import ExcelApp
from .exceptions import OffipyError
from .ppt import PptApp
from .word import WordApp


def _call(app: str, op: str, request_id: str | None = None, **kwargs):
    """转成 8890 server 调用；失败抛 RuntimeError 让模型看到原因。"""
    try:
        resp = request(app, op, request_id=request_id, **kwargs)
    except OffipyError as e:
        raise RuntimeError(str(e)) from None
    if not resp.get("ok"):
        raise RuntimeError(resp.get("error", "未知错误"))
    # OperationResult 契约：优先 data，旧 server 无 data 时回退 result
    return resp["data"] if "data" in resp else resp.get("result")


def _invoke(app: str, op: str, request_id: str | None = None, **kwargs):
    """统一返回封装：void op 返回 "ok (op)"，有值 op 结构化透传（int/list/str 等）。

    COM 对象结果在 server 侧 _serialize 成 null，落到这里就是 None → 返回
    确认串；真实有值的 op（总页数、表格数、文件列表、单元格值…）原样透传。
    """
    result = _call(app, op, request_id=request_id, **kwargs)
    return result if result is not None else f"ok ({op})"


server = MCPServer(
    name="offipy",
    title="offipy Office COM 自动化",
    description=(
        "会话式驱动真实 Microsoft Word/Excel/PowerPoint。只读操作作用于当前"
        "活动文档；改动/导出操作要求 doc_id、follow_active 或 expected_target "
        "显式绑定目标，防焦点漂移误操作。窗口实时可见、状态跨调用保持，产物"
        "是原生 Office 文件。"
    ),
    version=__version__,
    log_level="WARNING",
)

# ------------------------------------------------------------------ 注册

# MCP 工具名后缀覆盖：op 名 → 对外工具名，保持既有命名习惯
# （new_book → excel_new_workbook，read_doc_text → word_read_document_text 等）。
_TOOL_SUFFIX = {
    ("excel", "new_book"): "new_workbook",
    ("excel", "open_book"): "open_workbook",
    ("word", "new_doc"): "new_document",
    ("word", "open_doc"): "open_document",
    ("word", "read_doc_text"): "read_document_text",
    ("ppt", "new_pres"): "new_presentation",
    ("ppt", "open_pres"): "open_presentation",
}

# schema.returns → 动态函数返回注解（structured output 检测用）
_RETURN_ANNOTATION = {
    "void": str,
    "int": int,
    "str": str,
    "bool": bool,
    "list": list,
    "dict": dict,
    "any": object,
}

_APP_CLASSES = {"excel": ExcelApp, "word": WordApp, "ppt": PptApp}


def _tool_name(app: str, op: str) -> str:
    return f"{app}_{_TOOL_SUFFIX.get((app, op), op)}"


def _build_tool(app: str, op: str) -> None:
    """按 schema + App 方法签名生成一个 MCP 工具并注册到 server。

    参数签名去 self 后照搬 App 方法（必填/默认值/注解以方法为权威）；调用时
    补齐未传参数的默认值，保证 direct 调用与 MCP 框架路径行为一致。
    """
    fn_name = _tool_name(app, op)
    spec = schema.spec(app, op)
    if spec is None:
        return  # schema 之外的 op 不注册（遍历 schema 驱动，正常不会到这）
    method = getattr(_APP_CLASSES[app], op)
    params = [p for p in inspect.signature(method).parameters.values() if p.name != "self"]
    defaults = {p.name: p.default for p in params if p.default is not inspect.Parameter.empty}
    # 传输层参数（P0-1/P0-3）：破坏性 op 额外暴露 expected_target（JSON 对象绑定）
    # 与 follow_active（显式跟随活动文档），随 args 透传给 server dispatch。
    if schema.supports_expected_target(app, op):
        params = params + [
            inspect.Parameter(
                "expected_target",
                inspect.Parameter.KEYWORD_ONLY,
                annotation=dict,
                default=None,
            ),
            inspect.Parameter(
                "follow_active",
                inspect.Parameter.KEYWORD_ONLY,
                annotation=bool,
                default=False,
            ),
        ]
        defaults["expected_target"] = None
        defaults["follow_active"] = False
    return_ann = _RETURN_ANNOTATION.get(spec.returns, object)

    def tool_fn(ctx: Context | None = None, **kwargs: Any) -> Any:
        args = dict(defaults)
        args.update(kwargs)
        # 传输层参数（P0-1/P0-3）未给时不下发（None/False 语义等于缺省），
        # 避免请求 payload 噪音；给定了则透传，server dispatch 消费。
        if not args.get("expected_target"):
            args.pop("expected_target", None)
        if not args.get("follow_active"):
            args.pop("follow_active", None)
        # P1-4：MCP 框架自动注入 ctx（不出现在 tool schema），取其 request_id
        # 作幂等标识透传 server；无 ctx（测试直调）则 None。
        rid = getattr(ctx, "request_id", None) if ctx is not None else None
        return _invoke(app, op, request_id=rid, **args)

    fn = cast(Any, tool_fn)
    fn.__name__ = fn_name
    fn.__qualname__ = fn_name
    fn.__doc__ = spec.description
    fn.__signature__ = inspect.Signature(params, return_annotation=return_ann)

    if spec.readonly:
        annotations = ToolAnnotations(read_only_hint=True)
    else:
        # P0-3：导出 op 写文件系统，同样标记破坏性提示（requires_target ∪ destructive）
        annotations = ToolAnnotations(destructive_hint=spec.destructive or spec.requires_target)

    server.tool(name=fn_name, description=spec.description, annotations=annotations)(tool_fn)
    globals()[fn_name] = tool_fn


def _register_tools() -> None:
    """schema 驱动注册：除 quit（对整个会话太危险）外全部暴露。"""
    for app in schema.apps():
        for op in schema.ops(app):
            if op == "quit":
                continue
            _build_tool(app, op)


_register_tools()


def main():
    # stdio 传输：MCP 客户端（Claude Desktop 等）以子进程方式拉起并接管 stdin/stdout
    server.run()


if __name__ == "__main__":
    main()
