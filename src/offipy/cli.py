"""offipy CLI：把每个 Office 原子操作映射为一条子命令。

用法：
    office excel new_book
    office excel set_cell --sheet 1 --cell A1 --value 100
    office word new_doc
    office word write_line --text "你好"
    office ppt new_pres
    office ppt add_slide --layout 2
    office deck make --html examples/decks/starter/deck.html --out out/deck.pptx
    office quit excel
    office mcp                      # 启动 MCP stdio server（Claude Desktop 接入）

首次调用会自动在后台拉起常驻 server；之后所有操作都打到同一进程，
窗口持续可见、会话状态跨调用保持。
"""

import argparse
import json
import os

from .client import call, convert_value, ensure_server


def _parse_kwargs(tokens):
    kwargs = {}
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok.startswith("--"):
            key = tok[2:]
            if i + 1 < len(tokens) and not tokens[i + 1].startswith("--"):
                kwargs[key] = convert_value(tokens[i + 1])
                i += 2
            else:
                kwargs[key] = True
                i += 1
        else:
            kwargs[tok] = True
            i += 1
    return kwargs


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="office", description="offipy CLI")
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
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    if args.app == "mcp":
        from .mcp_server import main as mcp_main

        mcp_main()
        return
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
                    "用法: office deck make --html <deck.html> "
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
                    "用法: office deck outline --input <outline.md> "
                    "[--theme <name>] [--out <deck.html>]"
                )
            with open(md_path, encoding="utf-8") as f:
                outline = parse_outline(f.read())
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
    result = call(args.app, args.op, **kw)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
