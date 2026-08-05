"""offipy CLI：把每个 Office 原子操作映射为一条子命令。

用法：
    offipy excel new_book
    offipy excel set_cell --sheet 1 --cell A1 --value 100
    offipy word new_doc
    offipy word write_line --text "你好"
    offipy ppt new_pres
    offipy ppt add_slide --layout 2
    offipy deck make --html examples/decks/starter/deck.html --out out/deck.pptx
    offipy quit excel
    offipy mcp                      # 启动 MCP stdio server（Claude Desktop 接入）
    offipy check [--json]           # 环境就绪诊断（Python/依赖/Office/浏览器/server）

首次调用会自动在后台拉起常驻 server；之后所有操作都打到同一进程，
窗口持续可见、会话状态跨调用保持。

参数系统（P1）：
- 普通 `--key value`：值自动转 bool/int/float，无法转时回退字符串
- 重复 `--key`：多次出现聚合为 list
- `--payload '<json>'` / `--json '<json>'`：复杂嵌套参数以 JSON 透传，
  覆盖同名 kwargs
- 未知 `--key`（不在该 op 签名内）→ unrecognized arguments，exit 2
"""

import argparse
import inspect
import json
import os
import sys
import types
import typing

from . import excel, ppt, schema, word
from .client import call, ensure_server, server_status, set_port, stop_server
from .exceptions import OffipyError

_APP_CLASSES = {
    "excel": excel.ExcelApp,
    "word": word.WordApp,
    "ppt": ppt.PptApp,
}


def _parse_kwargs(tokens):
    """--key value 解析：token 值保留原始字符串，重复 key 聚合为 list。

    类型转换不再在这里做（弃全局猜测）：值按 schema 声明类型由 _coerce_kwargs
    统一转换，避免 "00123" 丢前导零、"true" 被误当 bool。--payload/--json 仍
    JSON 透传（结构化值覆盖同名 kwargs）。
    """
    kwargs = {}
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok in ("--payload", "--json"):
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                try:
                    payload = json.loads(tokens[i + 1])
                except (ValueError, TypeError) as e:
                    raise SystemExit(f"--{tok[2:]} JSON 解析失败: {e}") from None
                if not isinstance(payload, dict):
                    raise SystemExit(f"--{tok[2:]} 必须是 JSON 对象")
                kwargs.update(payload)
                i += 2
            else:
                kwargs[tok[2:]] = True
                i += 1
        elif tok in ("--expected-target", "--expected_target"):
            # P0-1/P0-3 传输层参数：目标绑定（JSON 对象），解析后存进 expected_target。
            # 值必须是对象（doc_id/name/path 之一），具体校验在 server 侧统一做。
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                try:
                    value = json.loads(tokens[i + 1])
                except (ValueError, TypeError) as e:
                    raise SystemExit(f"{tok} JSON 解析失败: {e}") from None
                if not isinstance(value, dict):
                    raise SystemExit(
                        f'{tok} 必须是 JSON 对象（如 --expected-target \'{{"doc_id":"book1"}}\'）'
                    )
                kwargs["expected_target"] = value
                i += 2
            else:
                raise SystemExit(f"{tok} 需要一个 JSON 对象值")
        elif tok in ("--follow-active", "--follow_active"):
            # P0-1/P0-3 传输层参数：显式声明跟随当前活动文档（裸用为 True，
            # 也接受 `--follow-active false`）。
            if (
                i + 1 < len(tokens)
                and not tokens[i + 1].startswith("--")
                and tokens[i + 1].strip().lower() in _BOOL_TOKENS
            ):
                kwargs["follow_active"] = _BOOL_TOKENS[tokens[i + 1].strip().lower()]
                i += 2
            else:
                kwargs["follow_active"] = True
                i += 1
        elif tok.startswith("--"):
            key = tok[2:]
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                value = tokens[i + 1]
                if key in kwargs:
                    existing = kwargs[key]
                    if isinstance(existing, list):
                        existing.append(value)
                    else:
                        kwargs[key] = [existing, value]
                else:
                    kwargs[key] = value
                i += 2
            else:
                kwargs[key] = True
                i += 1
        else:
            kwargs[tok] = True
            i += 1
    return kwargs


_BOOL_TOKENS = {
    "true": True,
    "false": False,
    "1": True,
    "0": False,
    "on": True,
    "off": False,
    "yes": True,
    "no": False,
}


class _BoolAction(argparse.Action):
    """--flag / --flag <true|false> 两用布尔：裸用为 True，显式值按 token 解析。

    根除 bool("false") is True 陷阱：显式传 false 必须是 False。
    """

    def __init__(self, option_strings, dest, **kwargs):
        kwargs.setdefault("nargs", "?")
        kwargs.setdefault("const", True)
        kwargs.setdefault("default", False)
        super().__init__(option_strings, dest, **kwargs)

    def __call__(self, parser, namespace, values, option_string=None):
        if values is None:
            setattr(namespace, self.dest, True)
            return
        token = str(values).strip().lower()
        if token not in _BOOL_TOKENS:
            raise argparse.ArgumentError(
                self, f"布尔值（true/false/1/0/on/off/yes/no），收到: {values!r}"
            )
        setattr(namespace, self.dest, _BOOL_TOKENS[token])


def _unwrap_optional(ann):
    """Optional[X] / X | None → X；其余注解原样返回。"""
    origin = typing.get_origin(ann)
    if origin is not typing.Union and origin is not getattr(types, "UnionType", None):
        return ann
    rest = [a for a in typing.get_args(ann) if a is not type(None)]
    if len(rest) == 1:
        return rest[0]
    return ann


def _coerce_fail(key: str, expected: str, value) -> None:
    print(f"offipy: error: --{key} 需要{expected}，收到: {value!r}", file=sys.stderr)
    raise SystemExit(2)


def _coerce_value(key: str, ann, value):
    """按单个参数注解转换；失败 stderr + exit 2（argparse 语义）。"""
    base = _unwrap_optional(ann)
    if base is str or base is inspect.Parameter.empty:
        return value
    if base is int:
        try:
            return int(value)
        except (TypeError, ValueError):
            _coerce_fail(key, "整数", value)
    if base is float:
        try:
            return float(value)
        except (TypeError, ValueError):
            _coerce_fail(key, "数字", value)
    if base is bool:
        # bool("false") 是 True 的陷阱必须避开：只认显式布尔 token
        if isinstance(value, bool):
            return value
        token = str(value).strip().lower()
        if token in _BOOL_TOKENS:
            return _BOOL_TOKENS[token]
        _coerce_fail(key, "布尔值（true/false/1/0/on/off/yes/no）", value)
    if base is list or typing.get_origin(base) is list:
        # 裸 list 与 list[T] 都按列表处理：标量包一层，列表原样
        return value if isinstance(value, list) else [value]
    return value


def _param_hints(app: str, op: str) -> dict[str, object]:
    """参数名→类型：优先 schema 声明（P1-2 单一来源），缺失时回退方法签名。

    未知 op / 无法检视签名 → 空表（参数原样透传，交 server 侧校验）。
    """
    sp = schema.spec(app, op)
    if sp is not None and sp.params:
        return dict(sp.params)
    cls = _APP_CLASSES.get(app)
    method = getattr(cls, op, None) if cls else None
    if method is None:
        return {}
    try:
        sig = inspect.signature(method)
    except (TypeError, ValueError):
        return {}
    return {name: param.annotation for name, param in sig.parameters.items() if name != "self"}


def _coerce_kwargs(app: str, op: str, kwargs: dict) -> dict:
    """按 schema 声明类型转换参数（P1-2，取代逐方法签名猜测）。

    无注解/Any/str 保持字符串；int/float/bool/list 按声明类型转换；未知
    op 跳过（交给 server 侧处理）。
    """
    hints = _param_hints(app, op)
    coerced = dict(kwargs)
    for key, value in kwargs.items():
        ann = hints.get(key)
        if ann is None:
            continue
        coerced[key] = _coerce_value(key, ann, value)
    return coerced


def _validate_kwargs(app: str, op: str, kwargs: dict) -> None:
    """未知 --key（不在 schema/方法参数内）→ stderr 报错并 exit 2（argparse 语义）。"""
    sp = schema.spec(app, op)
    if sp is not None and sp.params:
        known = set(sp.params)
    else:
        cls = _APP_CLASSES.get(app)
        method = getattr(cls, op, None) if cls else None
        if method is None:
            return  # 未知 op 交由 server 侧报错
        try:
            params = inspect.signature(method).parameters.values()
        except (TypeError, ValueError):
            return
        if any(p.kind is inspect.Parameter.VAR_KEYWORD for p in params):
            return
        known = {
            p.name
            for p in params
            if p.name != "self"
            and p.kind
            in (
                inspect.Parameter.POSITIONAL_OR_KEYWORD,
                inspect.Parameter.KEYWORD_ONLY,
                inspect.Parameter.POSITIONAL_ONLY,
            )
        }
    # 传输层参数（P0-1/P0-3）：破坏性 op 可显式绑定 expected_target / follow_active。
    # 非破坏性 op 上放行也无妨——server 侧对 expected_target 严格拒绝、follow_active
    # 静默忽略，语义一致。
    if schema.supports_expected_target(app, op):
        known |= {"expected_target", "follow_active"}
    for key in kwargs:
        if key not in known:
            print(f"offipy: error: {app} {op}: unrecognized arguments: --{key}", file=sys.stderr)
            raise SystemExit(2)


_TYPE_LABEL = {str: "str", int: "int", float: "num", bool: "bool"}


def _required_params(app: str, op: str) -> frozenset[str]:
    """App 方法签名中无默认值的参数（必填）——schema 未独立声明必填，以此派生。"""
    cls = _APP_CLASSES.get(app)
    method = getattr(cls, op, None) if cls else None
    if method is None:
        return frozenset()
    try:
        sig = inspect.signature(method)
    except (TypeError, ValueError):
        return frozenset()
    return frozenset(
        p.name
        for p in sig.parameters.values()
        if p.name != "self"
        and p.default is inspect.Parameter.empty
        and p.kind
        in (
            inspect.Parameter.POSITIONAL_OR_KEYWORD,
            inspect.Parameter.KEYWORD_ONLY,
        )
    )


def _validate_required(app: str, op: str, kwargs: dict) -> None:
    """必填参数缺失 → 调用前预校验：stderr 报错 + 用法，exit 2。

    不再等拉起 server / 碰 COM 后才炸；--payload 注入的键也算已提供。
    """
    missing = _required_params(app, op) - set(kwargs)
    if not missing:
        return
    sp = schema.spec(app, op)
    hints = dict(sp.params) if sp is not None and sp.params else {}
    shown = ", ".join(
        f"--{m}" + (f" <{_TYPE_LABEL[hints[m]]}>" if hints.get(m) in _TYPE_LABEL else "")
        for m in sorted(missing)
    )
    print(
        f"offipy: error: {app} {op}: 缺少必填参数 {shown}\n"
        f"  用法: offipy {app} {op} --<参数> <值> ...",
        file=sys.stderr,
    )
    raise SystemExit(2)


def _has_doc_id_param(app: str, op: str) -> bool:
    """op 是否接受 doc_id 参数（schema 声明或 App 方法签名）。quit 等无目标 op 返回 False。"""
    sp = schema.spec(app, op)
    if sp is not None and sp.params:
        return "doc_id" in sp.params
    cls = _APP_CLASSES.get(app)
    method = getattr(cls, op, None) if cls else None
    if method is None:
        return False
    try:
        sig = inspect.signature(method)
    except (TypeError, ValueError):
        return False
    return "doc_id" in sig.parameters


def _validate_destructive_target(app: str, op: str, kwargs: dict) -> None:
    """破坏性 op 缺目标（doc_id/expected_target/follow_active 均无）→ 提前友好报错。

    P0-3 doc_id 权威：与 server 的 InvalidArgumentError 同语义，但 CLI 先给
    usage 提示（exit 2），不等到拉起 server/碰 COM 后才炸。
    """
    if op == "quit" or not schema.supports_expected_target(app, op):
        return
    if not _has_doc_id_param(app, op):
        return  # 无 doc_id 参数的 op（理论上只有 quit）不需要目标
    if kwargs.get("doc_id") not in (None, ""):
        return
    if "expected_target" in kwargs:
        return
    if kwargs.get("follow_active"):
        return
    print(
        f"offipy: error: {app} {op}: 破坏性操作必须显式指定目标文档\n"
        f"  用法: offipy {app} {op} --doc_id <id> ...\n"
        f"        或 --expected-target '<json>' 绑定目标（如 '{{\"doc_id\":\"book1\"}}'）\n"
        f"        或 --follow-active 跟随当前活动文档",
        file=sys.stderr,
    )
    raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="offipy", description="offipy CLI")
    # P2-2 多实例：--port 指向指定端口的 server（env OFFIPY_SERVER_PORT 亦可）
    p.add_argument("--port", type=int, help="连指定端口的 server 实例（默认 8890）")
    sub = p.add_subparsers(dest="app", required=True)
    for app in ("excel", "word", "ppt"):
        sp = sub.add_parser(app)
        sp.add_argument("op")
        # REMAINDER：原样捕获 --key value 形式的任意 kwargs
        sp.add_argument("kwargs", nargs=argparse.REMAINDER)
    deck = sub.add_parser("deck")
    deck.add_argument("action", choices=["make", "outline"])
    # 专用选项（P0-4）：布尔用 _BoolAction，`--overwrite false` 不再是
    # bool("false")→True 的坑；未知 --key 由 argparse 直接 exit 2。
    deck.add_argument("--html", help="HTML 幻灯片源文件（make 必填）")
    deck.add_argument("--out", help="输出路径（make: .pptx；outline: .html）")
    deck.add_argument("--no-open", action="store_true", help="渲染后不打开实况演示（make）")
    deck.add_argument("--feedback", help="导出 PNG 反馈目录（make）")
    deck.add_argument("--theme", help="注入内置主题名（make/outline）")
    deck.add_argument("--layouts", action=_BoolAction, help="注入 data-layout 布局 CSS（make）")
    deck.add_argument("--overwrite", action=_BoolAction, help="覆盖已存在的 .pptx（make）")
    deck.add_argument("--input", help="大纲 markdown 源文件（outline）")
    deck.add_argument("--md", help="大纲 markdown 源文件别名（outline）")
    # 参数名避开顶层 subparsers 的 dest "app"，否则 argparse 会用子解析器的
    # 值覆盖 args.app（例如 quit ppt → app 被覆盖成 "ppt"）。
    q = sub.add_parser("quit")
    q.add_argument("target", choices=["excel", "word", "ppt"])
    sub.add_parser("mcp", help="启动 MCP stdio server（Claude Desktop 等接入）")
    ck = sub.add_parser("check", help="检查环境就绪（Python/依赖/Office/浏览器/server）")
    ck.add_argument("--json", action="store_true", help="输出 JSON")
    ck.add_argument(
        "--profile",
        choices=["core", "office", "deck", "mcp", "all"],
        default=None,
        help="只检查指定 profile（core 为基线，office/deck/mcp/all 各叠加对应分组）",
    )
    srv = sub.add_parser("server", help="管理常驻 server（status/stop/restart）")
    srv.add_argument(
        "action",
        nargs="?",
        default="status",
        choices=["status", "stop", "restart"],
    )
    # 子解析器的 --port 用 SUPPRESS：未给时不覆盖顶层值（否则默认 None 会把
    # `offipy --port 8891 server status` 的端口冲掉）。
    srv.add_argument(
        "--port", type=int, default=argparse.SUPPRESS, help="目标 server 端口（默认 8890）"
    )
    lg = sub.add_parser("log", help="读取操作日志（oplog.jsonl，P2-3）")
    lg.add_argument("--tail", type=int, help="只显示末尾 N 条")
    lg.add_argument(
        "--port", type=int, default=argparse.SUPPRESS, help="目标 server 端口（默认 8890）"
    )
    return p


def _main(argv=None):
    args = build_parser().parse_args(argv)
    if getattr(args, "port", None):
        set_port(args.port)  # P2-2 多实例：后续调用指向该端口
    if args.app == "mcp":
        try:
            from .mcp_server import main as mcp_main
        except ImportError as exc:
            print(
                "offipy: error: offipy mcp 需要 mcp 扩展：pip install offipy[mcp]\n"
                f"  缺失依赖: {exc}",
                file=sys.stderr,
            )
            return 2
        mcp_main()
        return
    if args.app == "check":
        from .envcheck import main as check_main

        return check_main(json_output=args.json, profile=getattr(args, "profile", None))
    if args.app == "quit":
        ensure_server()
        call(args.target, "quit")
        print("quit ok")
        return
    if args.app == "server":
        if args.action == "status":
            # 只读探测（P0-3）：未运行不拉起，直接报状态
            st = server_status()
            if st is None:
                print("server 未在运行")
            else:
                print(json.dumps(st, ensure_ascii=False))
        elif args.action == "stop":
            print(stop_server())
        else:  # restart
            print(stop_server())
            ensure_server()
            print("server 已重启")
        return
    if args.app == "log":
        from . import oplog

        oplog.configure(getattr(args, "port", None) or 8890)
        entries = oplog.read(tail=args.tail)
        if not entries:
            print("（暂无操作日志）")
            return
        for e in entries:
            print(json.dumps(e, ensure_ascii=False))
        return
    if args.app == "deck":
        if args.action == "make":
            from .deck import make as deck_make

            if not args.html:
                raise SystemExit(
                    "用法: offipy deck make --html <deck.html> "
                    "[--out <x.pptx>] [--no-open] [--feedback <dir>] "
                    "[--theme <name>] [--layouts] [--overwrite]"
                )
            pptx = deck_make(
                args.html,
                out=args.out,
                open_live_flag=not args.no_open,
                feedback_dir=args.feedback,
                theme=args.theme,
                apply_layouts=args.layouts,
                overwrite=args.overwrite,
            )
            print(json.dumps({"pptx": pptx}, ensure_ascii=False))
        else:  # outline
            from .outline import parse_outline, to_deck_html

            md_path = args.input or args.md
            if not md_path:
                raise SystemExit(
                    "用法: offipy deck outline --input <outline.md> "
                    "[--theme <name>] [--out <deck.html>]"
                )
            try:
                with open(md_path, encoding="utf-8") as f:
                    outline = parse_outline(f.read())
            except FileNotFoundError:
                raise SystemExit(f"找不到文件: {md_path}") from None
            except ValueError as e:
                raise SystemExit(f"大纲格式错误: {e}") from None
            html = to_deck_html(outline, theme=args.theme)
            if args.out:
                with open(args.out, "w", encoding="utf-8") as f:
                    f.write(html)
                print(json.dumps({"html": os.path.abspath(args.out)}, ensure_ascii=False))
            else:
                print(outline.to_json())
        return
    kw = _parse_kwargs(args.kwargs)
    _validate_kwargs(args.app, args.op, kw)
    _validate_required(args.app, args.op, kw)
    _validate_destructive_target(args.app, args.op, kw)
    kw = _coerce_kwargs(args.app, args.op, kw)
    result = call(args.app, args.op, **kw)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))


def main(argv=None):
    """offipy CLI 入口：库异常转 stderr + exit 1；SystemExit/argparse 原样放行。"""
    try:
        return _main(argv)
    except OffipyError as e:
        print(f"offipy: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
