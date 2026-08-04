"""offipy 客户端：常驻 server 的 HTTP 调用封装。"""

import json
import os
import subprocess
import sys
import time
import urllib.request

from .exceptions import RemoteCallError, ServerStartError
from .paths import user_data_dir

HOST = "127.0.0.1"
PORT = 8890
SERVER_MOD = "offipy.server"
_TOKEN_FILENAME = "token"
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


def _token() -> str | None:
    """读取 server 落盘的鉴权 token；不存在/不可读（旧 server）返回 None。"""
    try:
        token = (user_data_dir() / _TOKEN_FILENAME).read_text(encoding="utf-8").strip()
    except OSError:
        return None
    return token or None


def ensure_server():
    if _ping():
        return
    # 日志落盘用户数据目录：server 崩溃时能查根因（首次 gencache 生成类型库可能耗时）
    logpath = user_data_dir() / ".offipy.log"
    logpath.parent.mkdir(parents=True, exist_ok=True)
    popen_kwargs = {}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    with open(logpath, "a", encoding="utf-8") as logfile:
        subprocess.Popen(
            [sys.executable, "-m", SERVER_MOD, "--port", str(PORT)],
            stdout=logfile,
            stderr=logfile,
            **popen_kwargs,
        )
    for _ in range(600):  # 最多等 60 秒（首次 gencache 可能较慢）
        if _ping():
            return
        time.sleep(0.1)
    raise ServerStartError(f"无法启动 offipy server，请查看 {logpath}")


# 这些参数是文件/目录路径，必须在 client 侧按调用方 CWD 绝对化——
# server 是独立进程，其 CWD 与用户调 CLI 时所在目录无关
_PATH_KEYS = ("path", "out", "out_dir", "html", "pptx")


def request(app: str, op: str, **args) -> dict:
    """发一次调用，返回完整响应 dict（{ok, result, error, trace}），不做失败处理。

    供 MCP server 等调用方自行决定如何处理失败；CLI 仍走 call()。
    """
    ensure_server()
    for k in _PATH_KEYS:
        if k in args and isinstance(args[k], str):
            args[k] = os.path.abspath(args[k])
    data = json.dumps({"app": app, "op": op, "args": args}).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    token = _token()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(_URL + "/call", data=data, headers=headers)
    with _OPENER.open(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8"))


def call(app: str, op: str, **args):
    resp = request(app, op, **args)
    if not resp.get("ok"):
        msg = f"[{app}::{op}] 失败: {resp.get('error')}"
        if resp.get("trace"):
            msg += "\n" + "\n".join("  " + line for line in resp["trace"])
        raise RemoteCallError(msg)
    return resp.get("result")


def convert_value(v: str):
    s = v.strip()
    low = s.lower()
    if low == "true":
        return True
    if low == "false":
        return False
    if low in ("none", "null"):
        return None
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return v
