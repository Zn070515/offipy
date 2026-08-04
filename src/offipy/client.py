"""offipy 客户端：常驻 server 的 HTTP 调用封装。"""

import contextlib
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request

from .exceptions import RemoteCallError, ServerStartError
from .paths import user_data_dir

HOST = "127.0.0.1"
PORT = 8890
SERVER_MOD = "offipy.server"
_TOKEN_FILENAME = "token"
_URL = f"http://{HOST}:{PORT}"
# 部分机器会把系统代理写进注册表且 ProxyOverride 为空，连 127.0.0.1 回环请求
# 也会被劫持给代理（返回 502）。本地回环必须强制直连；真正出站的请求才该走代理。
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _ping() -> bool:
    try:
        with _OPENER.open(f"{_URL}/ping", timeout=1):
            return True
    except Exception:
        return False


def _auth_headers() -> dict:
    headers = {"Content-Type": "application/json"}
    token = _token()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


def _server_ok() -> bool:
    """真实握手：/status 鉴权通过且协议匹配才算 server 就绪（比裸 ping 可靠）。"""
    try:
        req = urllib.request.Request(f"{_URL}/status", headers=_auth_headers())
        with _OPENER.open(req, timeout=2) as r:
            if r.status != 200:
                return False
            data = json.loads(r.read().decode("utf-8"))
            return (
                bool(data.get("ok")) and data.get("result", {}).get("protocol") == "offipy-http/v1"
            )
    except Exception:
        return False


def _kill_pid(pid: int) -> None:
    """终止指定进程；Windows 用 taskkill 强杀，其他平台 os.kill。"""
    if sys.platform == "win32":
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    else:
        os.kill(pid, 9)


def _find_server_pid() -> int | None:
    """定位占用 8890 端口的 server 进程。

    netstat 优先（它是真正的端口持有者，杀了才能释放端口）；pid 文件兜底
    （进程刚拉起、尚未 LISTENING 时 netstat 查不到）。若倒过来以 pid 文件
    优先，一旦 pid 文件陈旧（指向已死或未持端口的进程），杀错目标、端口
    不释放，新 server 绑定失败，ensure_server 会一直握手失败到超时。
    """
    try:
        out = subprocess.run(["netstat", "-ano"], capture_output=True, text=True).stdout
    except OSError:
        out = ""
    for line in out.splitlines():
        if "LISTENING" not in line.upper():
            continue
        parts = line.split()
        # 本地地址形如 127.0.0.1:8890 或 [::]:8890，精确 endswith 防 :88900 误配
        if any(p.endswith(f":{PORT}") for p in parts[:2]):
            pid = parts[-1]
            if pid.isdigit():
                return int(pid)
    try:
        raw = (user_data_dir() / "server.pid").read_text(encoding="utf-8").strip()
        if raw.isdigit():
            return int(raw)
    except OSError:
        pass
    return None


def server_status() -> dict:
    """GET /status 返回字段 dict；失败抛 RemoteCallError。"""
    ensure_server()
    try:
        req = urllib.request.Request(f"{_URL}/status", headers=_auth_headers())
        with _OPENER.open(req, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
            if not data.get("ok"):
                raise RemoteCallError(data.get("error", "status 未知错误"))
            return data["result"]
    except urllib.error.HTTPError as e:
        raise RemoteCallError(f"status 失败: HTTP {e.code}") from e
    except urllib.error.URLError as e:
        raise RemoteCallError(f"status 连接失败: {e.reason}") from e


def stop_server() -> bool:
    """终止 8890 上的 offipy server；无运行进程返回 False。"""
    pid = _find_server_pid()
    if pid is None:
        return False
    try:
        _kill_pid(pid)
    except OSError:
        return False
    return True


def _token() -> str | None:
    """env 优先，其次 server 落盘的持久 token；都拿不到返回 None。"""
    env = os.environ.get("OFFIPY_SERVER_TOKEN")
    if env and env.strip():
        return env.strip()
    try:
        token = (user_data_dir() / _TOKEN_FILENAME).read_text(encoding="utf-8").strip()
    except OSError:
        return None
    return token or None


def ensure_server():
    """确保 8890 上跑着可用的 offipy server。

    用 /status 握手（非裸 ping）确认协议与鉴权都对。端口上有进程但握手失败
    （旧 server / token 不匹配）→ 定位其 pid 干掉重启；定位不到则报错并提示
    手动 `offipy server stop`。拉起新进程后写 server.pid 供 stop 定位。
    """
    if _server_ok():
        return
    if _ping():
        pid = _find_server_pid()
        if pid is None:
            raise ServerStartError(
                "端口 8890 已有不匹配的 offipy server，且无法定位其进程；"
                "请先运行 `offipy server stop` 或手动关闭后重试"
            )
        _kill_pid(pid)
    # 日志落盘用户数据目录：server 崩溃时能查根因（首次 gencache 生成类型库可能耗时）
    logpath = user_data_dir() / ".offipy.log"
    logpath.parent.mkdir(parents=True, exist_ok=True)
    pid_file = user_data_dir() / "server.pid"
    pid_file.parent.mkdir(parents=True, exist_ok=True)
    popen_kwargs = {}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    with open(logpath, "a", encoding="utf-8") as logfile:
        proc = subprocess.Popen(
            [sys.executable, "-m", SERVER_MOD, "--port", str(PORT)],
            stdout=logfile,
            stderr=logfile,
            **popen_kwargs,
        )
    # pid 文件写不了不致命：netstat 可兜底定位
    with contextlib.suppress(OSError):
        pid_file.write_text(str(proc.pid), encoding="utf-8")
    for _ in range(600):  # 最多等 60 秒（首次 gencache 可能较慢）
        if _server_ok():
            return
        time.sleep(0.1)
    raise ServerStartError(f"无法启动 offipy server，请查看 {logpath}")


# 这些参数是文件/目录路径，必须在 client 侧按调用方 CWD 绝对化——
# server 是独立进程，其 CWD 与用户调 CLI 时所在目录无关
_PATH_KEYS = ("path", "out", "out_dir", "html", "pptx")


def request(app: str, op: str, **args) -> dict:
    """发一次调用，返回完整响应 dict（{ok, result, error, trace}），不做失败处理。

    应用层失败（ok:false）仍以 dict 返回，供 MCP server 等调用方自行处理；
    HTTP 传输层失败（400/401/413/415/500/超时/连不上/坏 JSON）统一转成
    RemoteCallError，不再裸抛 HTTPError。
    """
    ensure_server()
    for k in _PATH_KEYS:
        if k in args and isinstance(args[k], str):
            args[k] = os.path.abspath(args[k])
    data = json.dumps({"app": app, "op": op, "args": args}).encode("utf-8")
    req = urllib.request.Request(_URL + "/call", data=data, headers=_auth_headers())
    try:
        with _OPENER.open(req, timeout=120) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode("utf-8"))
            detail = body.get("error") or e.reason
        except (ValueError, OSError):
            detail = f"HTTP {e.code}: {e.reason}"
        raise RemoteCallError(f"[{app}::{op}] 失败: {detail}") from e
    except urllib.error.URLError as e:
        raise RemoteCallError(f"[{app}::{op}] 连接失败: {e.reason}") from e
    except TimeoutError as e:
        raise RemoteCallError(f"[{app}::{op}] 调用超时: {e}") from e
    except json.JSONDecodeError as e:
        raise RemoteCallError(f"[{app}::{op}] 响应非 JSON: {e}") from e


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
