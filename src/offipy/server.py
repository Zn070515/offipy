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

import json
import os
import platform
import secrets
import time
import traceback
from http.server import BaseHTTPRequestHandler, HTTPServer

from . import __version__
from .excel import ExcelApp
from .exceptions import ServerStartError
from .paths import user_data_dir
from .ppt import PptApp
from .word import WordApp

DEFAULT_PORT = 8890
_MAX_BODY = 16 * 1024 * 1024  # 请求体上限 16MB
_TOKEN_FILENAME = "token"
_STARTED_AT = time.time()

_APPS: dict[str, object] = {}
_APPS_CLASSES = {
    "excel": ExcelApp,
    "word": WordApp,
    "ppt": PptApp,
}
# 操作白名单：显式注册表，新增 RPC 必须手动登记（勿用 dir() 反射——
# 会把 active_doc/active_book/active_pres 等会话内部方法暴露成远程可调）。
_OPS = {
    "excel": frozenset(
        {
            "new_book",
            "open_book",
            "close_book",
            "save",
            "save_pdf",
            "add_sheet",
            "set_cell",
            "get_cell",
            "set_range",
            "set_col_width",
            "format_cell",
            "merge_cells",
            "unmerge_cells",
            "set_border",
            "freeze_panes",
            "page_setup",
            "add_conditional_format",
            "set_row_height",
            "set_number_format",
            "autofit",
            "quit",
        }
    ),
    "word": frozenset(
        {
            "new_doc",
            "open_doc",
            "close_doc",
            "save",
            "save_pdf",
            "write",
            "write_line",
            "add_heading",
            "add_table",
            "set_table_cell",
            "format_text",
            "format_paragraph",
            "set_header_text",
            "set_footer_text",
            "add_page_number",
            "page_setup",
            "insert_toc",
            "update_toc",
            "add_list",
            "merge_table_cells",
            "set_table_border",
            "set_table_col_width",
            "set_table_row_height",
            "autofit_table",
            "find_replace",
            "insert_image",
            "insert_page_break",
            "quit",
        }
    ),
    "ppt": frozenset(
        {
            "new_pres",
            "open_pres",
            "save",
            "save_pdf",
            "export_slides",
            "add_slide",
            "set_title",
            "set_body",
            "set_notes",
            "add_textbox",
            "add_picture",
            "quit",
        }
    ),
}

# 运行时鉴权 token；serve() 启动时装载，Handler 在请求时读取
_TOKEN = ""


def _load_token() -> str:
    """env 优先，其次持久文件。

    env token 存在则直接返回（无需落盘）；否则必须落盘供 client 读取，
    写失败抛 ServerStartError——server 不应以 client 读不到 token 的
    假活状态启动。
    """
    env = os.environ.get("OFFIPY_SERVER_TOKEN")
    token = (env or "").strip()
    if token:
        return token
    token_file = user_data_dir() / _TOKEN_FILENAME
    if token_file.exists():
        token = token_file.read_text(encoding="utf-8").strip()
    if not token:
        token = secrets.token_urlsafe(32)
    try:
        token_file.parent.mkdir(parents=True, exist_ok=True)
        token_file.write_text(token, encoding="utf-8")
    except OSError as e:
        raise ServerStartError(f"无法写入 token 文件 {token_file}: {e}") from e
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


def dispatch(app, op: str, args: dict):
    if op.startswith("_"):
        raise PermissionError(f"不允许调用私有操作: {op}")
    if op != "quit" and not _alive(app):
        # 调用前主动保活检测：COM 引用已失效（用户关窗/Office 退出）则重建。
        # quit 例外：目标就是退出，app 已死时应直接成功，不反拉起新实例。
        app = _rebuild(app)
    method = getattr(app, op, None)
    if method is None:
        raise AttributeError(f"未知操作: {op}")
    try:
        return method(**args)
    except _com_error() as e:
        if getattr(e, "hresult", None) not in _DISCONNECTED_HRS:
            raise
        if op == "quit":
            return None
        # 调用中对象断连：重建实例重试一次
        return getattr(_rebuild(app), op)(**args)


def _check_auth(handler) -> bool:
    # 恒定时间比较，避免按长度早期短路泄露 token 信息
    return secrets.compare_digest(handler.headers.get("Authorization", ""), f"Bearer {_TOKEN}")


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
                        "protocol": "offipy-http/v1",
                        "pid": os.getpid(),
                        "python": platform.python_version(),
                        "started_at": _STARTED_AT,
                    },
                }
            )
        else:
            self._reply({"ok": False, "error": "not found"}, status=404)

    def do_POST(self):
        if not _check_auth(self):
            return self._reply({"ok": False, "error": "unauthorized"}, status=401)
        ctype = (self.headers.get("Content-Type", "") or "").split(";")[0].strip().lower()
        if ctype != "application/json":
            return self._reply(
                {"ok": False, "error": "Content-Type 必须是 application/json"}, status=415
            )
        try:
            n = int(self.headers.get("Content-Length", 0) or 0)
        except (TypeError, ValueError):
            return self._reply({"ok": False, "error": "Content-Length 无效"}, status=400)
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
            return self._reply({"ok": False, "error": f"未知应用: {app_name}"}, status=400)
        if op not in allowed:
            return self._reply({"ok": False, "error": f"未知操作: {app_name}::{op}"}, status=400)
        # 单线程 HTTPServer 下 handler 在主线程执行；COM 在 serve() 里
        # 初始化一次并常驻，这里不再 CoInit/CoUninit（反复拆建套间会
        # 使跨请求的 COM 对象失效，报“对象没有连接到服务器”）
        try:
            app = get_app(app_name)
            result = dispatch(app, op, body.get("args", {}))
            self._reply({"ok": True, "result": _serialize(result)})
        except Exception as e:
            tb = traceback.format_exc().strip().splitlines()
            self._reply(
                {
                    "ok": False,
                    "error": f"{type(e).__name__}: {e}",
                    "trace": tb[-3:],
                },
                status=500,
            )

    def _reply(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        pass


class Server(HTTPServer):
    # 单线程：COM 对象绑定创建线程的套间，多线程跨套间访问会出错。
    # 自用场景一次一个操作，单线程足够且最稳。
    # 禁止端口复用：防止多个 server 实例抢绑同一端口导致请求漂移。
    allow_reuse_address = False


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
    _TOKEN = _load_token()
    print(f"offipy server listening on http://{host}:{port}", flush=True)
    # 主线程初始化 COM 一次并保持整个服务生命周期
    import pythoncom  # 惰性：保 import offipy.server 跨平台可跑

    pythoncom.CoInitialize()
    try:
        Server((host, port), Handler).serve_forever()
    finally:
        pythoncom.CoUninitialize()


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
