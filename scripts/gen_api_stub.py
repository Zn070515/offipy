"""生成 src/offipy/api.pyi 类型 stub（供 mypy/IDE，勿手改）。

从 schema.OPS（op 元数据/返回类型/参数类型）＋ App 类方法签名（参数顺序/
默认值，唯一权威）生成 6 个 facade 类的显式方法签名。破坏性 op 的传输层参数
按入口区分（与 MCP/CLI 同策略）：
- 本地直连 Excel()/Word()/Ppt()：App 方法的 @destructive 守卫额外接受 follow_active
- 远程 Remote*()：额外接受 expected_target + follow_active + request_id

运行：`uv run python scripts/gen_api_stub.py`（纯标准库）。
tests/test_api_stub.py 兜底断言每个 schema op 都出现在 stub 且再生成无漂移。
"""

from __future__ import annotations

import inspect
from pathlib import Path
from typing import Any

from offipy import schema
from offipy.excel import ExcelApp
from offipy.ppt import PptApp
from offipy.word import WordApp

STUB = Path(__file__).resolve().parent.parent / "src" / "offipy" / "api.pyi"

_APP_CLASSES = {"excel": ExcelApp, "word": WordApp, "ppt": PptApp}
_DIRECT = {"excel": "Excel", "word": "Word", "ppt": "Ppt"}
_REMOTE = {"excel": "RemoteExcel", "word": "RemoteWord", "ppt": "RemotePpt"}

_TYPE_NAMES = {
    str: "str",
    bool: "bool",
    int: "int",
    float: "float",
    list: "list",
    dict: "dict",
    type(None): "None",
}
_RETURNS = {
    "void": "None",
    "str": "str",
    "int": "int",
    "bool": "bool",
    "str|null": "str | None",
    "list": "list",
    "list[SlideTextRecord]": "list[SlideTextRecord]",
    "dict": "dict",
    "any": "Any",
}

_HEADER = """\
\"\"\"offipy 高层 API 类型 stub（由 scripts/gen_api_stub.py 生成，勿手改）。

本地直连 Excel()/Word()/Ppt() 与远程 RemoteExcel()/RemoteWord()/RemotePpt()
的显式 op 签名，供 mypy/IDE。破坏性 op 传输层参数：direct 只走 follow_active，
remote 额外走 expected_target + request_id（与 MCP/CLI 同策略）。doc_id 缺省走
当前活动文档。
direct facade 绑定创建它的线程（STA COM），非线程安全；跨线程各自建 facade
时用 offipy.direct.com_apartment() 包一层（线程各自 CoInitialize/Uninitialize）。
\"\"\"

from typing import Any

from offipy.models import SlideTextRecord
"""


def _ann(t: Any) -> str:
    if t is Any:
        return "Any"
    return _TYPE_NAMES.get(t, "Any")


def _returns(r: str) -> str:
    return _RETURNS.get(r, "Any")


def _method_sig(app: str, op: str, remote: bool) -> str:
    """生成单个 op 的方法签名行（含传输层参数）。"""
    method = getattr(_APP_CLASSES[app], op)
    spec = schema.spec(app, op)
    assert spec is not None, f"{app}.{op} 未在 schema 登记"
    sig = inspect.signature(method)
    parts = []
    has_star = False
    has_doc_id = False
    for name, p in sig.parameters.items():
        if name == "self" or p.kind in (
            inspect.Parameter.VAR_POSITIONAL,
            inspect.Parameter.VAR_KEYWORD,
        ):
            continue
        # 忠实保留 App 方法签名的 keyword-only 分隔符（如 read_slide_texts 的
        # include_empty/recursive/doc_id），IDE/mypy 补全才能对上调用约定
        if p.kind == inspect.Parameter.KEYWORD_ONLY and not has_star:
            parts.append("*")
            has_star = True
        if name == "doc_id":
            has_doc_id = True
            if p.default is inspect.Parameter.empty:
                parts.append("doc_id: str | None")
            else:
                parts.append("doc_id: str | None = None")
            continue
        ann = spec.params.get(name)
        ann_str = _ann(ann) if ann is not None else _ann(p.annotation)
        if p.default is inspect.Parameter.empty:
            parts.append(f"{name}: {ann_str}")
        else:
            # App 方法用「默认 None」表达可省略参数（隐式 Optional）；
            # stub 显式补 `| None`，避免 mypy 报 Incompatible default。
            if p.default is None and ann_str not in ("Any", "None", "str | None"):
                ann_str = f"{ann_str} | None"
            parts.append(f"{name}: {ann_str} = {p.default!r}")
    if has_doc_id and schema.supports_expected_target(app, op):
        if not has_star:
            parts.append("*")
            has_star = True
        if remote:
            parts.append("expected_target: dict | None = None")
        parts.append("follow_active: bool = False")
    if remote:
        # P1-4：远程 facade 额外暴露 request_id（幂等标识，keyword-only）
        if not has_star:
            parts.append("*")
            has_star = True
        parts.append("request_id: str | None = None")
    params_str = f", {', '.join(parts)}" if parts else ""
    return f"    def {op}(self{params_str}) -> {_returns(spec.returns)}: ..."


def _render_class(cls: str, app: str, remote: bool) -> list[str]:
    lines = [f"class {cls}:"]
    if remote:
        lines.append("    def __init__(self, base_url: str | None = None) -> None: ...")
    else:
        lines.append(
            "    def __init__(self, visible: bool = True, "
            "modify_existing_visibility: bool = False) -> None: ..."
        )
    lines.append(f"    def __enter__(self) -> {cls}: ...")
    lines.append("    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> bool: ...")
    lines.append("")
    # 用 dict（插入序）而非 schema.ops() 的 frozenset——后者跨进程迭代顺序
    # 不稳定（hash 随机化），会让漂移测试偶发失败
    for op in schema.OPS[app]:
        lines.append(_method_sig(app, op, remote))
    lines.append("")
    return lines


def render() -> str:
    """拼装完整 stub 文本（不写盘，供 main 与测试复用）。"""
    lines = [_HEADER.rstrip(), ""]
    for app in schema.apps():
        lines += _render_class(_DIRECT[app], app, remote=False)
        lines += _render_class(_REMOTE[app], app, remote=True)
    lines.append("def op(app: str, op_name: str, **kw: Any) -> Any: ...")
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    STUB.write_text(render(), encoding="utf-8")
    print(f"generated {STUB}")


if __name__ == "__main__":
    main()
