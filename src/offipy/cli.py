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

from . import excel, ppt, word
from .client import call, convert_value, ensure_server
from .exceptions import OffipyError

_APP_CLASSES = {
    "excel": excel.ExcelApp,
    "word": word.WordApp,
    "ppt": ppt.PptApp,
}


def _parse_kwargs(tokens):
    """--key value 解析：重复 key 聚合为 list；--payload/--json 以 JSON 透传。"""
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
        elif tok.startswith("--"):
            key = tok[2:]
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                value = convert_value(tokens[i + 1])
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


def _validate_kwargs(app: str, op: str, kwargs: dict) -> None:
    """未知 --key（不在该 op 签名内）→ stderr 报错并 exit 2（argparse 语义）。"""
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
    for key in kwargs:
        if key not in known:
            print(f"offipy: error: {app} {op}: unrecognized arguments: --{key}", file=sys.stderr)
            raise SystemExit(2)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="offipy", description="offipy CLI")
    sub = p.add_subparsers(dest="app", required=True)
    for app in ("excel", "word", "ppt"):
        sp = sub.add_parser(app)
        sp.add_argument("op")
        # REMAINDER：原样捕获 --key value 形式的任意 kwargs
        sp.add_argument("kwargs", nargs=argparse.REMAINDER)
    deck = sub.add_parser("deck")
    deck.add_argument("action", choices=["make", "outline"])
    deck.add_argument("kwargs", nargs=argparse.REMAINDER)
    # 参数名避开顶层 subparsers 的 dest "app"，否则 argparse 会用子解析器的
    # 值覆盖 args.app（例如 quit ppt → app 被覆盖成 "ppt"）。
    q = sub.add_parser("quit")
    q.add_argument("target", choices=["excel", "word", "ppt"])
    sub.add_parser("mcp", help="启动 MCP stdio server（Claude Desktop 等接入）")
    ck = sub.add_parser("check", help="检查环境就绪（Python/依赖/Office/浏览器/server）")
    ck.add_argument("--json", action="store_true", help="输出 JSON")
    return p


def _main(argv=None):
    args = build_parser().parse_args(argv)
    if args.app == "mcp":
        from .mcp_server import main as mcp_main

        mcp_main()
        return
    if args.app == "check":
        from .envcheck import main as check_main

        return check_main(json_output=args.json)
    if args.app == "quit":
        ensure_server()
        call(args.target, "quit")
        print("quit ok")
        return
    if args.app == "deck":
        from .deck import make as deck_make

        kw = _parse_kwargs(args.kwargs)
        if args.action == "make":
            html = kw.pop("html", None)
            if not html:
                raise SystemExit(
                    "用法: offipy deck make --html <deck.html> "
                    "[--out <x.pptx>] [--no-open] [--feedback <dir>] "
                    "[--theme <name>] [--layouts]"
                )
            pptx = deck_make(
                html,
                out=kw.pop("out", None),
                open_live_flag=not (kw.pop("no-open", False) or kw.pop("no_open", False)),
                feedback_dir=kw.pop("feedback", None),
                theme=kw.pop("theme", None),
                apply_layouts=bool(kw.pop("layouts", False) or kw.pop("apply-layouts", False)),
            )
            print(json.dumps({"pptx": pptx}, ensure_ascii=False))
        elif args.action == "outline":
            from .outline import parse_outline, to_deck_html

            md_path = kw.pop("input", None) or kw.pop("md", None)
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
            html = to_deck_html(outline, theme=kw.pop("theme", None))
            out = kw.pop("out", None)
            if out:
                with open(out, "w", encoding="utf-8") as f:
                    f.write(html)
                print(json.dumps({"html": os.path.abspath(out)}, ensure_ascii=False))
            else:
                print(outline.to_json())
        return
    kw = _parse_kwargs(args.kwargs)
    _validate_kwargs(args.app, args.op, kw)
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
