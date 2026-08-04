"""office-kit CLI：把每个 Office 原子操作映射为一条子命令。

用法：
    office excel new_book
    office excel set_cell --sheet 1 --cell A1 --value 100
    office word new_doc
    office word write_line --text "你好"
    office ppt new_pres
    office ppt add_slide --layout 2
    office quit excel

首次调用会自动在后台拉起常驻 server；之后所有操作都打到同一进程，
窗口持续可见、会话状态跨调用保持。
"""

import argparse
import json
import os
import subprocess
import sys
import time
import urllib.request

HOST = "127.0.0.1"
PORT = 8890
SERVER_MOD = "office_kit.server"
_URL = f"http://{HOST}:{PORT}"
# 用户 VPN 在注册表写系统代理（ProxyServer=127.0.0.1:12334）且 ProxyOverride 为空，
# 会把本地 127.0.0.1:8890 回环请求也劫持给代理（返回 502）。
# 本地回环必须强制直连；真正出站的请求才该走代理。
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _ping() -> bool:
    try:
        with _OPENER.open(f"{_URL}/ping", timeout=1):
            return True
    except Exception:
        return False


def ensure_server():
    if _ping():
        return
    # 日志落盘：server 崩溃时能查根因（首次 gencache 生成类型库可能耗时）
    logpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".office-kit.log")
    with open(logpath, "a", encoding="utf-8") as logfile:
        subprocess.Popen(
            [sys.executable, "-m", SERVER_MOD, "--port", str(PORT)],
            stdout=logfile,
            stderr=logfile,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    for _ in range(600):  # 最多等 60 秒（首次 gencache 可能较慢）
        if _ping():
            return
        time.sleep(0.1)
    raise SystemExit("无法启动 office-kit server，请查看 .office-kit.log")


def call(app: str, op: str, **args):
    ensure_server()
    data = json.dumps({"app": app, "op": op, "args": args}).encode("utf-8")
    req = urllib.request.Request(
        _URL + "/call", data=data, headers={"Content-Type": "application/json"}
    )
    with _OPENER.open(req, timeout=120) as r:
        resp = json.loads(r.read().decode("utf-8"))
    if not resp.get("ok"):
        print(f"[{app}::{op}] 失败: {resp.get('error')}", file=sys.stderr)
        if resp.get("trace"):
            for line in resp["trace"]:
                print("  " + line, file=sys.stderr)
        raise SystemExit(1)
    return resp.get("result")


def _convert(v: str):
    if v in ("true", "True"):
        return True
    if v in ("false", "False"):
        return False
    if v in ("none", "None"):
        return None
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        pass
    return v


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="office", description="office-kit CLI")
    sub = p.add_subparsers(dest="app")
    for app in ("excel", "word", "ppt"):
        sp = sub.add_parser(app)
        sp.add_argument("op")
        # REMAINDER：原样捕获 --key value 形式的任意 kwargs，
        # 避免 argparse 把 --xx 当成未定义选项而报错
        sp.add_argument("kwargs", nargs=argparse.REMAINDER)
    # 注意：参数名避开顶层 subparsers 的 dest "app"，否则 argparse 会用
    # 子解析器的值覆盖 args.app（例如 quit ppt → app 被覆盖成 "ppt"）。
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
    kwargs = {}
    it = iter(args.kwargs)
    for tok in it:
        if tok.startswith("--"):
            kwargs[tok[2:]] = _convert(next(it))
        else:
            kwargs[tok] = True
    result = call(args.app, args.op, **kwargs)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
