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


def _probe() -> str:
    """探测 8890 上的 server 状态：ok / auth_fail / mismatch / down。

    - ok：/status 鉴权通过且协议匹配（我们可用的 server）
    - auth_fail：端口有进程回 401——进程可能是我们的 server，但 token 失配
    - mismatch：200 但协议不符——非 offipy 或旧版 server
    - down：连接失败/超时——无进程
    """
    try:
        req = urllib.request.Request(f"{_URL}/status", headers=_auth_headers())
        with _OPENER.open(req, timeout=2) as r:
            if r.status != 200:
                return "auth_fail" if r.status == 401 else "mismatch"
            data = json.loads(r.read().decode("utf-8"))
            if data.get("result", {}).get("protocol") == "offipy-http/v1":
                return "ok"
            return "mismatch"
    except urllib.error.HTTPError as e:
        # HTTPError 必须在此处理，否则会被下面的宽 except 当 "down"
        return "auth_fail" if e.code == 401 else "mismatch"
    except Exception:
        return "down"


def _server_ok() -> bool:
    """真实握手：/status 鉴权通过且协议匹配才算 server 就绪（比裸 ping 可靠）。"""
    return _probe() == "ok"


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


def server_ready() -> bool:
    """server 是否就绪（_probe == ok）。只读探测，不拉起。"""
    return _probe() == "ok"


def server_status() -> dict | None:
    """GET /status 返回字段 dict；server 未运行返回 None。

    只读探测：不调用 ensure_server()，绝不隐式拉起（P0-3）。
    """
    if _probe() != "ok":
        return None
    try:
        req = urllib.request.Request(f"{_URL}/status", headers=_auth_headers())
        with _OPENER.open(req, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
            if not data.get("ok"):
                return None
            return data["result"]
    except urllib.error.HTTPError as e:
        raise RemoteCallError(f"status 失败: HTTP {e.code}") from e
    except urllib.error.URLError as e:
        raise RemoteCallError(f"status 连接失败: {e.reason}") from e


def _pid_file_matches(pid: int) -> bool:
    """pid 文件记录的进程号与端口持有者一致，才认定是『我们的』server。"""
    try:
        raw = (user_data_dir() / "server.pid").read_text(encoding="utf-8").strip()
    except OSError:
        return False
    return raw.isdigit() and int(raw) == pid


def stop_server() -> str:
    """终止 8890 上的 offipy server，返回人工可读消息。

    进程所有权纪律（P0-1/P0-2）：能鉴权则走 /shutdown 优雅停机；token 失配
    或无法证明归属的进程一律不杀，只返回提示。
    """
    state = _probe()
    if state == "ok":
        # 身份由 token 证明，走鉴权 /shutdown 优雅停机（不依赖 pid 强杀）
        try:
            req = urllib.request.Request(f"{_URL}/shutdown", data=b"{}", headers=_auth_headers())
            with _OPENER.open(req, timeout=5):
                pass
        except Exception:
            pass  # /shutdown 404 等：旧版 server 无此端点，走下方回退
        # 优雅停机后确认进程退出；旧版 server 无 /shutdown 则回退 pid 强杀（需证明归属）
        for _ in range(100):
            if _probe() == "down":
                return "server 已停止"
            time.sleep(0.1)
        pid = _find_server_pid()
        if pid is not None and _pid_file_matches(pid):
            try:
                _kill_pid(pid)
            except OSError:
                return "终止失败（进程已退出或权限不足）"
            return "server 已停止"
        return "server 未能停止，请手动处理"
    if state == "auth_fail":
        return "server token 不匹配，拒绝终止；请设置正确的 OFFIPY_SERVER_TOKEN 后重试"
    if state == "mismatch":
        pid = _find_server_pid()
        if pid is not None and _pid_file_matches(pid):
            try:
                _kill_pid(pid)
            except OSError:
                return "终止失败（进程已退出或权限不足）"
            return "server 已停止"
        return "8890 端口被非 offipy 进程占用且无法确认归属，拒绝强杀；请手动处理"
    return "server 未在运行"


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

    用 /status 握手（非裸 ping）确认协议与鉴权都对。进程所有权纪律：
    - auth_fail（token 失配）→ 绝不杀，报错让用户修正 token；
    - mismatch（非 offipy / 旧版 server）→ 仅当 pid 文件证明归属才强杀重启，
      否则拒绝并提示手动处理（P0-1/P0-2）。
    拉起新进程后写 server.pid 供 stop 定位。
    """
    state = _probe()
    if state == "ok":
        return
    if state == "auth_fail":
        # P0-2：token 不匹配绝不杀 server——旧 client 连新 server 只报错，不误伤进程
        raise ServerStartError(
            "8890 端口的 offipy server token 不匹配（拒绝强杀）。"
            f"请设置正确的 OFFIPY_SERVER_TOKEN，或删除 {user_data_dir() / _TOKEN_FILENAME} "
            "后运行 `offipy server restart`"
        )
    if state == "mismatch":
        pid = _find_server_pid()
        if pid is None or not _pid_file_matches(pid):
            # P0-1：端口上有进程但无法证明是我们的 server → 拒绝强杀
            raise ServerStartError(
                "8890 端口已有不匹配的进程，且无法确认是 offipy server（拒绝强杀）。"
                "请运行 `offipy server stop` 或手动处理后重试"
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
