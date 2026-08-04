"""office-kit CLI：把每个 Office 原子操作映射为一条子命令。

用法：
    office excel new_book
    office excel set_cell --sheet 1 --cell A1 --value 100
    office word new_doc
    office word write_line --text "你好"
    office ppt new_pres
    office ppt add_slide --layout 2
    office deck make --html examples/decks/starter/deck.html --out out/deck.pptx
    office quit excel

首次调用会自动在后台拉起常驻 server；之后所有操作都打到同一进程，
窗口持续可见、会话状态跨调用保持。
"""

import argparse
import json

from .client import call, convert_value, ensure_server


def _parse_kwargs(tokens):
    kwargs = {}
    it = iter(tokens)
    for tok in it:
        if tok.startswith("--"):
            kwargs[tok[2:]] = convert_value(next(it))
        else:
            kwargs[tok] = True
    return kwargs


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="office", description="office-kit CLI")
    sub = p.add_subparsers(dest="app")
    for app in ("excel", "word", "ppt"):
        sp = sub.add_parser(app)
        sp.add_argument("op")
        # REMAINDER：原样捕获 --key value 形式的任意 kwargs
        sp.add_argument("kwargs", nargs=argparse.REMAINDER)
    deck = sub.add_parser("deck")
    deck.add_argument("action", choices=["make"])
    deck.add_argument("kwargs", nargs=argparse.REMAINDER)
    # 参数名避开顶层 subparsers 的 dest "app"，否则 argparse 会用子解析器的
    # 值覆盖 args.app（例如 quit ppt → app 被覆盖成 "ppt"）。
    q = sub.add_parser("quit")
    q.add_argument("target", choices=["excel", "word", "ppt"])
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
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
                    "[--out <x.pptx>] [--no-open] [--feedback <dir>]"
                )
            pptx = deck_make(
                html,
                out=kw.pop("out", None),
                open_live_flag=not (kw.pop("no-open", False) or kw.pop("no_open", False)),
                feedback_dir=kw.pop("feedback", None),
            )
            print(json.dumps({"pptx": pptx}, ensure_ascii=False))
        return
    kw = _parse_kwargs(args.kwargs)
    result = call(args.app, args.op, **kw)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
