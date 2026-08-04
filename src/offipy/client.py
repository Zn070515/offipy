"""offipy 客户端：常驻 server 的 HTTP 调用封装。"""

import json
import os
import subprocess
import sys
import time
import urllib.request

HOST = "127.0.0.1"
PORT = 8890
SERVER_MOD = "offipy.server"
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
    logpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".offipy.log")
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
    raise SystemExit("无法启动 offipy server，请查看 .offipy.log")


def request(app: str, op: str, **args) -> dict:
    """发一次调用，返回完整响应 dict（{ok, result, error, trace}），不做失败处理。

    供 MCP server 等调用方自行决定如何处理失败；CLI 仍走 call()。
    """
    ensure_server()
    data = json.dumps({"app": app, "op": op, "args": args}).encode("utf-8")
    req = urllib.request.Request(
        _URL + "/call", data=data, headers={"Content-Type": "application/json"}
    )
    with _OPENER.open(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8"))


def call(app: str, op: str, **args):
    resp = request(app, op, **args)
    if not resp.get("ok"):
        print(f"[{app}::{op}] 失败: {resp.get('error')}", file=sys.stderr)
        if resp.get("trace"):
            for line in resp["trace"]:
                print("  " + line, file=sys.stderr)
        raise SystemExit(1)
    return resp.get("result")


def convert_value(v: str):
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
