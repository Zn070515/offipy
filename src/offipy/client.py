"""offipy 客户端：常驻 server 的 HTTP 调用封装。"""

import hashlib
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from typing import NoReturn

from .exceptions import (
    ComOperationError,
    ConversionError,
    FileConflictError,
    InvalidArgumentError,
    OfficeUnavailableError,
    ProtocolError,
    RemoteCallError,
    ServerStartError,
    TargetNotFoundError,
    UnsupportedPlatformError,
)
from .paths import user_data_dir

# error_code（server 失败响应携带）→ 领域异常：RPC 错误与库异常一一对应（P1-4）
_ERROR_CODE_TO_EXC = {
    "invalid_argument": InvalidArgumentError,
    "target_not_found": TargetNotFoundError,
    "file_conflict": FileConflictError,
    "com_operation": ComOperationError,
    "protocol": ProtocolError,
    "server_start": ServerStartError,
    "conversion": ConversionError,
    "office_unavailable": OfficeUnavailableError,
    "remote_call": RemoteCallError,
    "unsupported_platform": UnsupportedPlatformError,
}

HOST = "127.0.0.1"
PORT = 8890  # 默认端口；多实例时用 set_port()/OFFIPY_SERVER_PORT 指向其他实例
PROTOCOL = "offipy-http/v1"  # 请求侧握手协议（P2-8），server 校验不匹配回 ProtocolError
# /call 超时与 server._CALL_TIMEOUT 对齐（§4 审计：旧值 120 vs server 600 错配）。
# 两边各持同名常量、取值一致，测试断言二者相等，防单边漂移。
_CALL_TIMEOUT = 600
# PID 归属时间窗口（P1-3）：pid 文件 started_at 与进程真实创建时间偏差超过
# 此值即视为 PID 已被复用（旧文件 + 新进程），拒认归属、不杀错目标。
_START_WINDOW = 60.0
SERVER_MOD = "offipy.server"
_TOKEN_FILENAME = "token"
# 部分机器会把系统代理写进注册表且 ProxyOverride 为空，连 127.0.0.1 回环请求
# 也会被劫持给代理（返回 502）。本地回环必须强制直连；真正出站的请求才该走代理。
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

_PORT = PORT


def port() -> int:
    """当前目标端口：OFFIPY_SERVER_PORT env 优先，其次 set_port() 设定。"""
    raw = os.environ.get("OFFIPY_SERVER_PORT")
    if raw and raw.strip().isdigit():
        return int(raw.strip())
    return _PORT


def set_port(p: int) -> None:
    """把后续调用指向指定端口的 server 实例（P2-2 多实例）。"""
    global _PORT
    _PORT = int(p)


def _base_url() -> str:
    return f"http://{HOST}:{port()}"


def _port_from_url(base_url: str | None) -> int:
    """从 base_url 解析端口（P1-1）；无 base_url → port()。https 缺省 443、http 缺省 80。"""
    if base_url is None:
        return port()
    m = re.search(r":(\d+)\b", base_url)
    if m:
        return int(m.group(1))
    if base_url.startswith("https://"):
        return 443
    if base_url.startswith("http://"):
        return 80
    return port()


def _token_path(p: int):
    # 默认端口沿用旧文件名（token），非默认端口按端口隔离（token-{port}）
    name = _TOKEN_FILENAME if p == PORT else f"{_TOKEN_FILENAME}-{p}"
    return user_data_dir() / name


def _pid_path(p: int):
    # 默认端口沿用旧文件名（server.pid），非默认端口按端口隔离（server-{port}.pid）
    name = "server.pid" if p == PORT else f"server-{p}.pid"
    return user_data_dir() / name


def _ping() -> bool:
    try:
        with _OPENER.open(f"{_base_url()}/ping", timeout=1):
            return True
    except Exception:
        return False


def _auth_headers(p: int) -> dict:
    headers = {"Content-Type": "application/json", "X-Offipy-Protocol": PROTOCOL}
    token = _token(p)
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
        req = urllib.request.Request(f"{_base_url()}/status", headers=_auth_headers(port()))
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
        if any(p.endswith(f":{port()}") for p in parts[:2]):
            pid = parts[-1]
            if pid.isdigit():
                return int(pid)
    try:
        raw = _pid_path(port()).read_text(encoding="utf-8").strip()
    except OSError:
        return None
    # 新格式 JSON {port,pid,...}；旧格式纯数字
    if raw.isdigit():
        return int(raw)
    try:
        data = json.loads(raw)
        pid = data.get("pid")
        if isinstance(pid, int):
            return pid
    except (ValueError, TypeError, AttributeError):
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
        req = urllib.request.Request(f"{_base_url()}/status", headers=_auth_headers(port()))
        with _OPENER.open(req, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
            if not data.get("ok"):
                return None
            return data["result"]
    except urllib.error.HTTPError as e:
        raise RemoteCallError(f"status 失败: HTTP {e.code}") from e
    except urllib.error.URLError as e:
        raise RemoteCallError(f"status 连接失败: {e.reason}") from e


def _process_start_time(pid: int) -> float | None:
    """进程创建时间（epoch 秒）；Windows 用 PowerShell 查询，失败返回 None。

    PID 复用防护（P1-3）：pid 文件记录的 started_at 必须与进程真实创建时间
    接近，否则说明该 pid 已被操作系统复用于别的进程——旧文件 + 新进程，
    归属不可信。非 Windows 返回 None（不校验）。
    """
    if sys.platform != "win32":
        return None
    try:
        out = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                (
                    f"(Get-Process -Id {pid} -ErrorAction SilentlyContinue)"
                    ".StartTime.ToUniversalTime().ToString('o')"
                ),
            ],
            capture_output=True,
            text=True,
            timeout=5,
        ).stdout.strip()
    except Exception:
        return None
    if not out:
        return None
    try:
        from datetime import datetime, timezone

        ts = datetime.fromisoformat(out.replace("Z", "+00:00"))
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
        return ts.timestamp()
    except (ValueError, TypeError):
        return None


def _pid_file_matches(pid: int) -> bool:
    """pid 文件 + 端口持有者 + token + 进程创建时间，多重要素才认定是『我们的』server。

    仅接受 JSON 格式 {port,pid,token_sha256,started_at}：pid 与端口持有者一致、
    token_sha256 与本地已知 token 一致、且 started_at 与进程真实创建时间
    相近才算归属（P0-1 强化 + P1-3 PID 复用防护）。旧格式纯数字文件无法
    证明 token 归属 → 一律拒绝（P0-2：不杀无法证明归属的进程）。
    """
    try:
        raw = _pid_path(port()).read_text(encoding="utf-8").strip()
    except OSError:
        return False
    if raw.isdigit():
        return False  # P0-2：纯数字 pid 文件不携带 token 证据，不认
    try:
        data = json.loads(raw)
    except (ValueError, TypeError):
        return False
    if not isinstance(data, dict) or data.get("pid") != pid:
        return False
    if data.get("port") is not None and data.get("port") != port():
        return False
    token = _token(port())
    if not token:
        return False
    if data.get("token_sha256") != hashlib.sha256(token.encode()).hexdigest():
        return False
    recorded_start = data.get("started_at")
    if isinstance(recorded_start, (int, float)) and recorded_start > 0:
        start = _process_start_time(pid)
        if start is not None and abs(start - recorded_start) > _START_WINDOW:
            return False  # PID 复用：文件是旧的，进程不是当年的 server
    return True


def stop_server() -> str:
    """终止 8890 上的 offipy server，返回人工可读消息。

    进程所有权纪律（P0-1/P0-2）：能鉴权则走 /shutdown 优雅停机；token 失配
    或无法证明归属的进程一律不杀，只返回提示。
    """
    state = _probe()
    if state == "ok":
        # 身份由 token 证明，走鉴权 /shutdown 优雅停机（不依赖 pid 强杀）
        try:
            req = urllib.request.Request(
                f"{_base_url()}/shutdown", data=b"{}", headers=_auth_headers(port())
            )
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


def _token(p: int) -> str | None:
    """env 优先，其次 server 落盘的持久 token（按端口隔离）；都拿不到返回 None。"""
    env = os.environ.get("OFFIPY_SERVER_TOKEN")
    if env and env.strip():
        return env.strip()
    try:
        token = _token_path(p).read_text(encoding="utf-8").strip()
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
            f"{port()} 端口的 offipy server token 不匹配（拒绝强杀）。"
            f"请设置正确的 OFFIPY_SERVER_TOKEN，或删除 {_token_path(port())} "
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
    logpath = user_data_dir() / f".offipy-{port()}.log"
    logpath.parent.mkdir(parents=True, exist_ok=True)
    pid_file = _pid_path(port())
    pid_file.parent.mkdir(parents=True, exist_ok=True)
    popen_kwargs = {}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    with open(logpath, "a", encoding="utf-8") as logfile:
        subprocess.Popen(
            [sys.executable, "-m", SERVER_MOD, "--port", str(port())],
            stdout=logfile,
            stderr=logfile,
            **popen_kwargs,
        )
    # pid 文件由 server 绑定成功后自写权威记录（P1-3），client 不抢写；
    # 定位兜底 netstat 始终可用。
    for _ in range(600):  # 最多等 60 秒（首次 gencache 可能较慢）
        if _server_ok():
            return
        time.sleep(0.1)
    raise ServerStartError(f"无法启动 offipy server，请查看 {logpath}")


# 这些参数是文件/目录路径，必须在 client 侧按调用方 CWD 绝对化——
# server 是独立进程，其 CWD 与用户调 CLI 时所在目录无关
_PATH_KEYS = ("path", "out", "out_dir", "html", "pptx")


def _raise_error(app: str, op: str, code, detail, hresult=None, trace=None) -> NoReturn:
    """按 server 错误码映射回领域异常；ComOperationError 透传 hresult（契约5）。"""
    msg = f"[{app}::{op}] 失败: {detail}"
    if trace:
        msg += "\n" + "\n".join("  " + line for line in trace)
    exc_cls = _ERROR_CODE_TO_EXC.get(code) if code else None
    if exc_cls is ComOperationError:
        hr = int(hresult, 16) if isinstance(hresult, str) else hresult
        raise ComOperationError(msg, hresult=hr)
    if exc_cls is not None:
        raise exc_cls(msg)
    raise RemoteCallError(msg)


def request(
    app: str, op: str, base_url: str | None = None, *, request_id: str | None = None, **args
) -> dict:
    """发一次调用并返回响应 dict；应用层失败抛对应 OffipyError。

    成功（server 返回 200 + ok:true）→ dict；应用层失败（server 以 HTTP 500
    返回 ok:false + error_code）→ 按 error_code 映射抛领域异常，com_operation
    透传 hresult 供断连识别；传输层失败（超时/连不上/坏 JSON/400/401/413/415
    等）→ RemoteCallError。base_url 缺省连本地 8890（P0-4 Remote* 共享 CLI/
    MCP 会话）；显式给出可指向其他 offipy server。

    request_id（P0-2 幂等方案 A）：缺省自动生成 uuid4；调用方超时重试应复用
    同一 request_id——server 对同 id 同 payload 合并/回放缓存，不重复执行。
    响应带 request_id 回显，可核对。
    """
    if base_url is None:
        ensure_server()
    for k in _PATH_KEYS:
        if k in args and isinstance(args[k], str):
            args[k] = os.path.abspath(args[k])
    # P1-2：expected_target.path 与其它路径参数一致，按调用方 CWD 绝对化——
    # server 的 path 规范化比较基于 abspath，相对写法在跨 CWD 时语义漂移。
    et = args.get("expected_target")
    if isinstance(et, dict) and isinstance(et.get("path"), str):
        args["expected_target"] = {**et, "path": os.path.abspath(et["path"])}
    if request_id is None:
        request_id = str(uuid.uuid4())
    # request_id 幂等标识（§4/方案 A）：client 重试带同一 id，server 命中缓存
    # 不再重执行；payload hash 绑定保证同 id 必须同 payload。
    data = json.dumps({"app": app, "op": op, "args": args, "request_id": request_id}).encode(
        "utf-8"
    )
    # P1-1：token 按 base_url 解析出的端口取（多实例隔离），绝不误用默认端口 token
    p = _port_from_url(base_url)
    req = urllib.request.Request(
        (base_url or _base_url()) + "/call", data=data, headers=_auth_headers(p)
    )
    try:
        with _OPENER.open(req, timeout=_CALL_TIMEOUT) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = None
        try:
            body = json.loads(e.read().decode("utf-8"))
            detail = body.get("error") or e.reason
            code = body.get("error_code")
        except (ValueError, OSError):
            detail = f"HTTP {e.code}: {e.reason}"
            code = None
        hresult = body.get("hresult") if isinstance(body, dict) else None
        _raise_error(app, op, code, detail, hresult=hresult)
    except urllib.error.URLError as e:
        raise RemoteCallError(f"[{app}::{op}] 连接失败: {e.reason}") from e
    except TimeoutError as e:
        raise RemoteCallError(f"[{app}::{op}] 调用超时: {e}") from e
    except json.JSONDecodeError as e:
        raise RemoteCallError(f"[{app}::{op}] 响应非 JSON: {e}") from e


def call(app: str, op: str, base_url: str | None = None, *, request_id: str | None = None, **args):
    resp = request(app, op, base_url=base_url, request_id=request_id, **args)
    if not resp.get("ok"):
        _raise_error(
            app,
            op,
            resp.get("error_code"),
            resp.get("error"),
            hresult=resp.get("hresult"),
            trace=resp.get("trace"),
        )
    # OperationResult 契约：优先 data，旧 server 无 data 时回退 result
    return resp["data"] if "data" in resp else resp.get("result")
