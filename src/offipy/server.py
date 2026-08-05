"""常驻会话服务：持有 Word/Excel/PPT 的 COM 引用，提供本地 HTTP 调用。

会话式驱动的关键：server 进程持有各 App 实例的引用不释放，Office 窗口
稳定存活；CLI 每次操作通过 HTTP 打到本服务，跨调用状态（打开的文档、
当前工作簿等）天然保持。

安全模型（P0-4）：
- 启动时生成/读取持久 token（env OFFIPY_SERVER_TOKEN 优先，否则
  user_data_dir()/token），/call 与 /status 须带 Authorization: Bearer <token>；
  校验失败仅 401，不杀 server——旧 client 连新 server 只报错，不误伤进程。
- 请求体限 16MB（超限 413）；Content-Type 必须 application/json（否则 415）。
- 操作白名单 _OPS：显式注册表，只放行逐条登记的方法（active_doc/active_book/
  active_pres 等会话内部方法一律不在其内），白名单外一律 400；
  dispatch 的 `_` 前缀 guard 保留为纵深防御。
"""

import contextlib
import hashlib
import json
import os
import platform
import queue
import secrets
import threading
import time
import traceback
import uuid
from collections import OrderedDict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any

from . import __version__, oplog, schema
from .excel import ExcelApp
from .exceptions import (
    ComOperationError,
    InvalidArgumentError,
    ServerStartError,
    TargetNotFoundError,
)
from .paths import user_data_dir
from .ppt import PptApp
from .result import OperationResult
from .word import WordApp

DEFAULT_PORT = 8890
_PROTOCOL = "offipy-http/v1"  # 请求侧握手协议（P2-8）
_MAX_BODY = 16 * 1024 * 1024  # 请求体上限 16MB
_MAX_RESPONSE = 64 * 1024 * 1024  # 响应上限 64MB（超限降级 500，不写大 payload）
_CALL_TIMEOUT = 600  # /call 入队后等 worker 结果的超时（安全兜底，op 本身由 client 超时）
# 有界资源（§4/§5）：队列与并发都设上限，满则快速失败（503），不做无限排队
_COM_QUEUE_MAX = 64  # COM worker 队列容量；满 → /call 立即 503 busy
_MAX_CONCURRENCY = 16  # 同时处理的 HTTP 连接上限；超出直接 503（防线程风暴）
# request_id 幂等缓存（§4）：最近 N 个已处理请求，重试命中直接返回缓存结果
_REQUEST_ID_MAX = 512  # LRU 上限
_REQUEST_ID_TTL = 600.0  # 缓存存活与 _CALL_TIMEOUT 同量级：超时重试窗口内有效
_TOKEN_FILENAME = "token"
_STARTED_AT = time.time()
# 本次 server 会话标识（P2-3）：随 /status 暴露并写入每条 oplog，供跨实例
# 区分/追查；多实例（P2-2）时天然扩展。
_SESSION_ID = str(uuid.uuid4())

# 目标身份资源类型（resource_id 的中间段）
_KINDS = {"excel": "book", "word": "doc", "ppt": "pres"}

_APPS: dict[str, Any] = {}
_APPS_CLASSES = {
    "excel": ExcelApp,
    "word": WordApp,
    "ppt": PptApp,
}
# 操作白名单：由 schema 单一来源派生（P1-2）。新增 RPC 只需在 schema.py
# 登记一条 OpSpec + 在 App 类实现方法，server 不再手工维护集合；
# `_` 前缀 guard 在 dispatch 保留为纵深防御（防 dir() 反射类内部方法）。
_OPS = {app: frozenset(schema.ops(app)) for app in schema.apps()}

# 破坏性 op 集合（会改动文档内容/状态）：由 schema 的 destructive 标志派生。
# expected_target 绑定只对这类 op 生效，防止用户焦点漂移后误改到非预期的文档。
_DESTRUCTIVE_OPS = {app: frozenset(schema.destructive_ops(app)) for app in schema.apps()}

# 单 COM worker（P1-1）：COM 对象只允许在创建它的线程里访问，所有 App
# 实例都绑定 worker 线程；HTTP handler 线程只入队/取回结果，慢 op 不阻塞
# /ping /status /shutdown。队列与 worker 均为模块级（与 _APPS 同级共享）。
# 有界队列（§4）：容量 _COM_QUEUE_MAX，满则 handler 立即 503，不无限阻塞。
_COM_QUEUE: "queue.Queue[tuple | None]" = queue.Queue(maxsize=_COM_QUEUE_MAX)
_WORKER: threading.Thread | None = None
_WORKER_LOCK = threading.Lock()

# request_id 幂等缓存：request_id → (monotonic 时间戳, 响应 dict)。LRU（过
# 期惰性清理 + 容量封顶）。handler 线程只读，GIL 下 get/set 原子够用；同一
# id 的并发首投仍可能双双入队（幂等针对「超时后顺序重试」，非分布式锁）。
_REQUEST_ID_CACHE: "OrderedDict[str, tuple[float, dict]]" = OrderedDict()

# /status 的目标身份缓存：worker 每次 op 后刷新（worker 线程持有 COM，探测
# 安全）；handler 线程只读快照，绝不因 status 探测拉起 Office 或触碰 COM。
_LAST_TARGETS: dict[str, dict | None] = {name: None for name in _APPS_CLASSES}

# 运行时鉴权 token；serve() 启动时装载，Handler 在请求时读取
_TOKEN = ""


def _token_path(port: int):
    # 默认端口沿用旧文件名（token），非默认端口按端口隔离（token-{port}）
    name = _TOKEN_FILENAME if port == DEFAULT_PORT else f"{_TOKEN_FILENAME}-{port}"
    return user_data_dir() / name


def _pid_path(port: int):
    # 默认端口沿用旧文件名（server.pid），非默认端口按端口隔离（server-{port}.pid）
    name = "server.pid" if port == DEFAULT_PORT else f"server-{port}.pid"
    return user_data_dir() / name


def _write_pid_file(port: int, token: str) -> None:
    """落盘 PID 文件（{port,pid,token_sha256,started_at}），供 client 归属验证。

    写失败不致命：client 可用 netstat 兜底定位；token_sha256 让 client 在
    强杀前证明归属（P0-1），对不上就不杀。
    """
    pid_file = _pid_path(port)
    with contextlib.suppress(OSError):
        pid_file.parent.mkdir(parents=True, exist_ok=True)
        pid_file.write_text(
            json.dumps(
                {
                    "port": port,
                    "pid": os.getpid(),
                    "token_sha256": hashlib.sha256(token.encode()).hexdigest(),
                    "started_at": _STARTED_AT,
                }
            ),
            encoding="utf-8",
        )


def _load_token(port: int) -> str:
    """env 优先，其次持久文件。

    env token 存在则直接返回（无需落盘）；否则必须落盘供 client 读取，
    写失败抛 ServerStartError——server 不应以 client 读不到 token 的
    假活状态启动。落盘 token 收紧为 0o600（§4：Windows 上 chmod 为 no-op，
    但 POSIX 下杜绝同机其他用户读取）。
    """
    env = os.environ.get("OFFIPY_SERVER_TOKEN")
    token = (env or "").strip()
    if token:
        return token
    token_file = _token_path(port)
    if token_file.exists():
        token = token_file.read_text(encoding="utf-8").strip()
        with contextlib.suppress(OSError):
            os.chmod(token_file, 0o600)  # 顺手收紧既有文件权限
    if not token:
        token = secrets.token_urlsafe(32)
    try:
        token_file.parent.mkdir(parents=True, exist_ok=True)
        token_file.write_text(token, encoding="utf-8")
    except OSError as e:
        raise ServerStartError(f"无法写入 token 文件 {token_file}: {e}") from e
    with contextlib.suppress(OSError):
        os.chmod(token_file, 0o600)
    return token


def get_app(name: str):
    cls = _APPS_CLASSES.get(name)
    if cls is None:
        raise ValueError(f"未知应用: {name}，可选 {list(_APPS_CLASSES)}")
    if name not in _APPS:
        _APPS[name] = cls()
    return _APPS[name]


def _serialize(v):
    if v is None or isinstance(v, (int, float, str, bool)):
        return v
    if isinstance(v, (list, tuple)):
        return [_serialize(x) for x in v]
    if isinstance(v, dict):
        return {str(k): _serialize(x) for k, x in v.items()}
    try:
        return v.isoformat()  # datetime 等
    except Exception:
        pass
    if hasattr(v, "_oleobj_"):  # COM 对象（new_book 等返回的 Workbook/Presentation）
        return None  # 序列化无意义，返回 null 而非 "<COMObject>"
    return str(v)


# 与 Office 进程断连/进程消失的 COM HRESULT
_DISCONNECTED_HRS = {
    0x80010111,  # RPC_E_DISCONNECTED（对象没有连接到服务器）
    0x80010001,  # RPC_E_SERVER_DIED
    0x80010008,  # RPC_E_INVALID_OBJECT（被调用的对象已与其客户端断开连接）
    0x800401FD,  # CO_E_OBJNOTCONNECTED
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
}


def _com_error():
    """惰性取 pywintypes.com_error；非 Windows 降级为 Exception，保 import 不炸。"""
    try:
        import pywintypes

        return pywintypes.com_error
    except ImportError:
        return Exception


def _alive(app) -> bool:
    """探测 app 持有的 COM 对象是否仍与 Office 进程保持连接。"""
    try:
        _ = app.app.Visible
        return True
    except (_com_error(), AttributeError):
        return False


def _rebuild(app):
    """丢弃失效的 App 实例并重建（复用新进程里已恢复的 Office）。"""
    name = next((k for k, v in _APPS.items() if v is app), None)
    if name is None:
        return app
    _APPS.pop(name, None)
    return get_app(name)


_TARGET_KEYS = ("doc_id", "name", "path")


def _resolve_expected_target(app, expected) -> str:
    """expected_target 绑定：resolve-once，返回校验通过后的 doc_id。

    P0-4/5：校验与执行用同一个 doc_id（校验时解析、注入方法调用参数），
    杜绝「校验 A 执行 B」；未知键/空对象直接拒绝，堵死旧 _target_matches({})
    恒真绕过。绑定失败抛 TargetNotFoundError / InvalidArgumentError。
    """
    if not isinstance(expected, dict):
        raise InvalidArgumentError(f"expected_target 必须是对象，收到 {type(expected).__name__}")
    unknown = [k for k in expected if k not in _TARGET_KEYS]
    if unknown:
        raise InvalidArgumentError(f"expected_target 含未知键: {unknown}（可选: doc_id/name/path）")
    if not any(k in expected for k in _TARGET_KEYS):
        raise InvalidArgumentError("expected_target 必须至少包含 doc_id/name/path 之一")
    target = app.get_target(doc_id=expected.get("doc_id"))
    if target is None:
        raise TargetNotFoundError("没有可绑定的目标文档")
    for key in ("name", "path"):
        if key in expected and target.get(key) != expected[key]:
            raise TargetNotFoundError(
                f"目标绑定失败: 期望 {key}={expected[key]!r}，实际 {target.get(key)!r}"
            )
    return target["doc_id"]


def _current_targets() -> dict[str, dict | None]:
    """当前会话各 App 的目标身份；未初始化的 App 不拉起，报 null。只读。"""
    targets: dict[str, dict | None] = {}
    for name in _APPS_CLASSES:
        app = _APPS.get(name)
        if app is None:
            targets[name] = None
            continue
        try:
            targets[name] = app.get_target()
        except Exception:
            targets[name] = None
    return targets


def _refresh_targets() -> None:
    """worker 线程内刷新 status 目标缓存（一次快照赋值，handler 读引用安全）。"""
    global _LAST_TARGETS
    with contextlib.suppress(Exception):
        _LAST_TARGETS = _current_targets()


def _resource_id(app_name: str, doc_id: str | None = None) -> str | None:
    """目标身份 → "app:kind:doc_id"；无目标/探测失败返回 None。

    P0-6：resource_id 用 doc_id（会话内稳定标识）而非 name（用户可改的文件名）。
    显式 doc_id 优先（dispatch 注入的绑定/调用方 doc_id）；否则取当前活动目标的 doc_id。
    """
    if doc_id:
        return f"{app_name}:{_KINDS.get(app_name, 'doc')}:{doc_id}"
    app = _APPS.get(app_name)
    if app is None:
        return None
    try:
        target = app.get_target()
    except Exception:
        return None
    if not target or not target.get("doc_id"):
        return None
    return f"{app_name}:{_KINDS.get(app_name, 'doc')}:{target['doc_id']}"


def _deprecation_warning(op_name: str) -> str | None:
    """已弃用 op 的 warning 文案（P2-9）；未弃用返回 None。"""
    app_name, _, op = op_name.partition(".")
    sp = schema.spec(app_name, op)
    if sp is not None and sp.deprecated:
        return f"{op_name} 已弃用（deprecated），将在未来版本移除"
    return None


def _success_result(op_name: str, result, doc_id: str | None = None) -> dict:
    """OperationResult 归一化成功响应；data 已 _serialize，另附 result 兼容别名。"""
    data = _serialize(result)
    app_name, _, _ = op_name.partition(".")
    res = OperationResult(
        ok=True,
        operation=op_name,
        resource_id=_resource_id(app_name, doc_id),
        message="ok",
        data=data,
    ).to_dict()
    res["result"] = data
    w = _deprecation_warning(op_name)
    if w:
        res["warning"] = w
    return res


def _error_result(op_name: str, e: Exception) -> dict:
    """失败响应：带 error_code（异常 code），client 据此映射回领域异常。"""
    tb = traceback.format_exc().strip().splitlines()
    res = {
        "ok": False,
        "operation": op_name,
        "resource_id": None,
        "error": f"{type(e).__name__}: {e}",
        "error_code": getattr(e, "code", "internal"),
        "trace": tb[-3:],
    }
    if isinstance(e, ComOperationError) and e.hresult is not None:
        res["hresult"] = hex(e.hresult)
    w = _deprecation_warning(op_name)
    if w:
        res["warning"] = w
    return res


def _append_oplog(app_name: str, op: str, kind: str, payload: dict, duration_ms: int) -> None:
    """每次 op 后落一条操作日志（P2-3）；失败静默，不拖垮请求。"""
    with contextlib.suppress(Exception):
        oplog.append(
            _SESSION_ID,
            app_name,
            op,
            ok=(kind == "ok"),
            error_code=payload.get("error_code") if kind == "error" else None,
            duration_ms=duration_ms,
            resource_id=payload.get("resource_id"),
        )


def _dedupe_hit(request_id: str) -> dict | None:
    """request_id 幂等命中：已处理返回缓存响应，未处理/缺省返回 None。

    顺带惰性清理过期项（TTL 窗口外的旧请求让位）。
    """
    if not request_id:
        return None
    now = time.monotonic()
    stale = [k for k, (ts, _) in _REQUEST_ID_CACHE.items() if now - ts > _REQUEST_ID_TTL]
    for k in stale:
        _REQUEST_ID_CACHE.pop(k, None)
    item = _REQUEST_ID_CACHE.get(request_id)
    if item is None:
        return None
    _REQUEST_ID_CACHE[request_id] = (now, item[1])  # 刷新 LRU 位置
    return item[1]


def _dedupe_store(request_id: str, payload: dict) -> None:
    """记录已处理 request_id 的响应；容量封顶（LRU 丢弃最旧）。"""
    if not request_id:
        return
    _REQUEST_ID_CACHE[request_id] = (time.monotonic(), payload)
    while len(_REQUEST_ID_CACHE) > _REQUEST_ID_MAX:
        _REQUEST_ID_CACHE.popitem(last=False)


def _ensure_worker() -> None:
    """懒启动单 COM worker（幂等）。COM 初始化移到 worker 线程（P1-1）。"""
    global _WORKER
    with _WORKER_LOCK:
        if _WORKER is not None and _WORKER.is_alive():
            return
        _WORKER = threading.Thread(target=_worker_loop, name="offipy-com-worker", daemon=True)
        _WORKER.start()


def _stop_worker() -> None:
    """停 worker：哨兵唤醒退出循环，join 有限等待（daemon 兜底不阻塞进程）。

    队列有界（§4）：哨兵放不进时最多等 _COM_QUEUE_MAX 满队被消费，超时放弃
    （daemon 线程随进程退出，不强行阻塞停机路径）。
    """
    global _WORKER
    with _WORKER_LOCK:
        w = _WORKER
        _WORKER = None
    if w is not None and w.is_alive():
        # 队列可能被长 op 占满：daemon 兜底，放不进哨兵就放弃，不无限等待
        with contextlib.suppress(queue.Full):
            _COM_QUEUE.put(None, timeout=5)
        w.join(timeout=5)


def _worker_loop() -> None:
    """COM worker 主循环：CoInitialize 绑定本线程套间，串行消费 op。

    任务元组 (app_name, op, args, resp_q)；None 哨兵退出。op 只在此线程
    执行，App 对象不会跨套间访问。
    """
    try:
        import pythoncom  # 惰性：保 import offipy.server 跨平台可跑

        pythoncom.CoInitialize()
        com_ready = True
    except Exception:
        # 非 Windows / 无 pywin32：worker 仍可消费队列（假 op/纯逻辑测试），
        # 真 COM op 会由各 App 自行报错，不因初始化失败拖垮整个队列。
        com_ready = False
    try:
        while True:
            item = _COM_QUEUE.get()
            if item is None:
                break
            app_name, op, args, resp_q = item
            op_name = f"{app_name}.{op}"
            t0 = time.monotonic()
            try:
                app = get_app(app_name)
                result = dispatch(app, op, args, app_name)
                # dispatch 可能注入绑定 doc_id（expected_target）或保留调用方 doc_id；
                # resource_id 用它定位——「操作的是谁，资源 id 就是谁」
                kind, payload = "ok", _success_result(op_name, result, doc_id=args.get("doc_id"))
            except Exception as e:
                kind, payload = "error", _error_result(op_name, e)
            finally:
                _refresh_targets()
            duration_ms = int((time.monotonic() - t0) * 1000)
            # 先落日志再回包：响应到达时记录已持久化，调用方 read 无竞态
            _append_oplog(app_name, op, kind, payload, duration_ms)
            resp_q.put((kind, payload))
    finally:
        if com_ready:
            pythoncom.CoUninitialize()


def dispatch(app, op: str, args: dict, app_name: str):
    if op.startswith("_"):
        raise PermissionError(f"不允许调用私有操作: {op}")
    if op != "quit" and not _alive(app):
        # 调用前主动保活检测：COM 引用已失效（用户关窗/Office 退出）则重建。
        # quit 例外：目标就是退出，app 已死时应直接成功，不反拉起新实例。
        app = _rebuild(app)
    expected = args.pop("expected_target", None)
    if expected is not None and op in _DESTRUCTIVE_OPS.get(app_name, frozenset()):
        # P0-4/5：破坏性 op 绑定目标——不跟随用户焦点。resolve-once：校验用
        # 解析出的 doc_id 直接注入方法参数，杜绝「校验 A 执行 B」；未知键/空
        # 对象在 _resolve_expected_target 内拒绝（旧 _target_matches({}) 恒真
        # 绕过已堵死）。只读 op 上的 expected_target 无意义：pop 掉忽略。
        args["doc_id"] = _resolve_expected_target(app, expected)
    method = getattr(app, op, None)
    if method is None:
        raise AttributeError(f"未知操作: {op}")
    try:
        return method(**args)
    except ComOperationError as e:
        # App 方法已由 guard_com 把 pywintypes.com_error 包成 ComOperationError；
        # 断连 HRESULT 才重建重试，其余 COM 失败原样上抛（error_code=com_operation）。
        if getattr(e, "hresult", None) not in _DISCONNECTED_HRS:
            raise
        if op == "quit":
            return None
        # 调用中对象断连：重建实例重试一次
        return getattr(_rebuild(app), op)(**args)


def _check_auth(handler) -> bool:
    # 恒定时间比较，避免按长度早期短路泄露 token 信息
    return secrets.compare_digest(handler.headers.get("Authorization", ""), f"Bearer {_TOKEN}")


def _encode_reply(obj, status: int = 200) -> tuple[int, bytes]:
    """序列化响应；超过响应上限降级为 500 错误，不向客户端写超大 payload。"""
    data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    if len(data) > _MAX_RESPONSE:
        return 500, json.dumps(
            {"ok": False, "error": f"响应超过 {_MAX_RESPONSE} 字节上限"},
            ensure_ascii=False,
        ).encode("utf-8")
    return status, data


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/ping":
            # 健康检查免鉴权：不暴露任何数据，供 client 探测存活
            self._reply({"ok": True, "result": "pong"})
        elif self.path == "/status":
            if not _check_auth(self):
                return self._reply({"ok": False, "error": "unauthorized"}, status=401)
            self._reply(
                {
                    "ok": True,
                    "result": {
                        "version": __version__,
                        "protocol": _PROTOCOL,
                        "session_id": _SESSION_ID,
                        "pid": os.getpid(),
                        "python": platform.python_version(),
                        "started_at": _STARTED_AT,
                        # 只读缓存快照：worker 线程持有 COM 探测，handler 不触碰
                        "targets": _LAST_TARGETS,
                    },
                }
            )
        else:
            self._reply({"ok": False, "error": "not found"}, status=404)

    def do_POST(self):
        if not _check_auth(self):
            return self._reply({"ok": False, "error": "unauthorized"}, status=401)
        if self.path in ("/call", "/shutdown") and (
            (self.headers.get("X-Offipy-Protocol") or "") != _PROTOCOL
        ):
            # P2-8 请求侧握手：/call /shutdown 必须带 protocol 头，不匹配回
            # ProtocolError——旧 client 连新 server 时协议协商失败，不静默错位。
            return self._reply(
                {
                    "ok": False,
                    "error": f"协议不匹配: 期望 {_PROTOCOL}",
                    "error_code": "protocol",
                },
                status=400,
            )
        if self.path == "/shutdown":
            # 鉴权过的优雅停机：回包后由独立线程触发 shutdown()——不能在 handler
            # 线程内直调（serve_forever 要等当前请求完成，会死锁）。
            self._reply({"ok": True, "result": "shutting down"})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return
        if self.path != "/call":
            return self._reply({"ok": False, "error": "not found"}, status=404)
        ctype = (self.headers.get("Content-Type", "") or "").split(";")[0].strip().lower()
        if ctype != "application/json":
            return self._reply(
                {"ok": False, "error": "Content-Type 必须是 application/json"}, status=415
            )
        try:
            n = int(self.headers.get("Content-Length", 0) or 0)
        except (TypeError, ValueError):
            return self._reply({"ok": False, "error": "Content-Length 无效"}, status=400)
        if n < 0:
            # read(负值) 会吞掉整个连接缓冲，必须显式拒绝
            return self._reply({"ok": False, "error": "Content-Length 不能为负"}, status=400)
        if n > _MAX_BODY:
            return self._reply(
                {"ok": False, "error": f"请求体超过 {_MAX_BODY} 字节上限"}, status=413
            )
        try:
            body = json.loads(self.rfile.read(n) or b"{}")
        except Exception as e:
            return self._reply({"ok": False, "error": f"JSON 解析失败: {e}"}, status=400)
        if not isinstance(body, dict):
            return self._reply({"ok": False, "error": "请求体必须是 JSON 对象"}, status=400)
        app_name = body.get("app")
        op = body.get("op")
        allowed = _OPS.get(app_name)
        if allowed is None:
            return self._reply(
                {"ok": False, "error": f"未知应用: {app_name}", "error_code": "invalid_argument"},
                status=400,
            )
        if op not in allowed:
            return self._reply(
                {
                    "ok": False,
                    "error": f"未知操作: {app_name}::{op}",
                    "error_code": "invalid_argument",
                },
                status=400,
            )
        # 校验（鉴权/路径/Content-Type/体积/白名单）留在 handler 线程 fail-fast；
        # COM op 入队给单 worker 串行执行，worker 结果经 per-request 队列取回。
        request_id = body.get("request_id")
        if not isinstance(request_id, str) or not request_id:
            request_id = None
        cached = _dedupe_hit(request_id)
        if cached is not None:
            # request_id 幂等（§4）：重复请求返回缓存结果，不重执行
            self._reply(cached, status=200 if cached.get("ok") else 500)
            return
        _ensure_worker()
        resp_q: queue.Queue[tuple] = queue.Queue(maxsize=1)
        try:
            _COM_QUEUE.put_nowait((app_name, op, body.get("args", {}), resp_q))
        except queue.Full:
            # 有界队列（§4）：满则立即 503，不让调用方无限排队
            return self._reply(
                {
                    "ok": False,
                    "error": "server 忙（COM 队列已满），请稍后重试",
                    "error_code": "busy",
                },
                status=503,
            )
        try:
            kind, payload = resp_q.get(timeout=_CALL_TIMEOUT)
        except queue.Empty:
            return self._reply(
                {
                    "ok": False,
                    "error": f"操作超时（worker 忙或卡住，>{_CALL_TIMEOUT}s）",
                    "error_code": "internal",
                },
                status=504,
            )
        _dedupe_store(request_id, payload)
        self._reply(payload, status=200 if kind == "ok" else 500)

    def _reply(self, obj, status=200):
        status, data = _encode_reply(obj, status)
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        pass


class Server(ThreadingHTTPServer):
    # 线程前端（P1-1）：每个 HTTP 请求独立线程，/ping /status /shutdown 不碰
    # COM、直处理；慢 COM op 在单 worker 里串行排队，不阻塞健康检查与停机。
    # 禁止端口复用：防止多个 server 实例抢绑同一端口导致请求漂移。
    # 有界线程（§4）：Semaphore 限并发 _MAX_CONCURRENCY，超出直接 503——防
    # 线程风暴拖垮进程（ThreadingMixIn 默认每连接一线程、无上限）。
    allow_reuse_address = False
    _SLOT = threading.BoundedSemaphore(_MAX_CONCURRENCY)

    def process_request(self, request, client_address):
        if not self._SLOT.acquire(blocking=False):
            # 并发已满：连接级直回 503，不让客户端挂到传输超时
            self._reply_503(request)
            return
        try:
            super().process_request(request, client_address)
        except Exception:
            self._SLOT.release()
            raise

    def process_request_thread(self, request, client_address):
        try:
            super().process_request_thread(request, client_address)
        finally:
            self._SLOT.release()

    def _reply_503(self, request) -> None:
        body = json.dumps(
            {"ok": False, "error": "server 忙（并发连接数已满）", "error_code": "busy"},
            ensure_ascii=False,
        ).encode("utf-8")
        resp = (
            b"HTTP/1.1 503 Service Unavailable\r\n"
            b"Content-Type: application/json\r\n"
            b"Content-Length: " + str(len(body)).encode("utf-8") + b"\r\n"
            b"Connection: close\r\n\r\n" + body
        )
        try:
            with contextlib.suppress(Exception):
                request.sendall(resp)
        finally:
            with contextlib.suppress(Exception):
                request.close()


_LOOPBACK_HOSTS = {"127.0.0.1", "localhost", "::1", ""}


def _validate_host(host: str, allow_remote: bool) -> None:
    """拒绝绑定非回环地址：token 挡不了端口扫描，回环是默认安全边界。"""
    if not allow_remote and host not in _LOOPBACK_HOSTS:
        raise ServerStartError(
            f"拒绝绑定非回环地址 {host!r}；如需远程访问请显式传 allow_remote=True"
        )


def serve(
    port: int = DEFAULT_PORT,
    host: str = "127.0.0.1",
    allow_remote: bool = False,
):
    global _TOKEN
    _validate_host(host, allow_remote)
    oplog.configure(port)  # P2-2 多实例：日志按端口隔离
    _TOKEN = _load_token(port)
    _write_pid_file(port, _TOKEN)  # §4 启动锁：port+pid+token_sha256 落盘，供归属验证
    print(f"offipy server listening on http://{host}:{port}", flush=True)
    # COM 初始化移到 worker 线程（P1-1）：HTTP 线程只入队，App 对象只被
    # worker 触碰，套间安全；/ping /status /shutdown 不碰 COM，不被排队。
    _ensure_worker()
    try:
        Server((host, port), Handler).serve_forever()
    finally:
        _stop_worker()


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument(
        "--unsafe-allow-remote",
        action="store_true",
        help="显式允许绑定非回环地址（有安全风险，仅测试/内网用）",
    )
    a = ap.parse_args()
    serve(a.port, a.host, allow_remote=a.unsafe_allow_remote)
