"""常驻会话服务：持有 Word/Excel/PPT 的 COM 引用，提供本地 HTTP 调用。

会话式驱动的关键：server 进程持有各 App 实例的引用不释放，Office 窗口
稳定存活；CLI 每次操作通过 HTTP 打到本服务，跨调用状态（打开的文档、
当前工作簿等）天然保持。
"""

import json
import traceback
from http.server import BaseHTTPRequestHandler, HTTPServer

import pythoncom
import pywintypes

from .excel import ExcelApp
from .ppt import PptApp
from .word import WordApp

DEFAULT_PORT = 8890
_APPS: dict[str, object] = {}
_APPS_CLASSES = {
    "excel": ExcelApp,
    "word": WordApp,
    "ppt": PptApp,
}


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
        return str(v)


# 与 Office 进程断连/进程消失的 COM HRESULT
_DISCONNECTED_HRS = {
    0x80010111,  # RPC_E_DISCONNECTED（对象没有连接到服务器）
    0x80010001,  # RPC_E_SERVER_DIED
    0x80010008,  # RPC_E_INVALID_OBJECT（被调用的对象已与其客户端断开连接）
    0x800401FD,  # CO_E_OBJNOTCONNECTED
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
}


def _alive(app) -> bool:
    """探测 app 持有的 COM 对象是否仍与 Office 进程保持连接。"""
    try:
        _ = app.app.Visible
        return True
    except (pywintypes.com_error, AttributeError):
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
    if not _alive(app):
        # 调用前主动保活检测：COM 引用已失效（用户关窗/Office 退出）则重建
        app = _rebuild(app)
    method = getattr(app, op, None)
    if method is None:
        raise AttributeError(f"未知操作: {op}")
    try:
        return method(**args)
    except pywintypes.com_error as e:
        if getattr(e, "hresult", None) not in _DISCONNECTED_HRS:
            raise
        # 调用中对象断连：重建实例重试一次
        return getattr(_rebuild(app), op)(**args)


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/ping":
            self._reply({"ok": True, "result": "pong"})
        else:
            self._reply({"ok": False, "error": "not found"})

    def do_POST(self):
        # 单线程 HTTPServer 下 handler 在主线程执行；COM 在 serve() 里
        # 初始化一次并常驻，这里不再 CoInit/CoUninit（反复拆建套间会
        # 使跨请求的 COM 对象失效，报“对象没有连接到服务器”）
        try:
            n = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(n) or b"{}")
            app = get_app(body["app"])
            result = dispatch(app, body["op"], body.get("args", {}))
            self._reply({"ok": True, "result": _serialize(result)})
        except Exception as e:
            tb = traceback.format_exc().strip().splitlines()
            self._reply(
                {
                    "ok": False,
                    "error": f"{type(e).__name__}: {e}",
                    "trace": tb[-3:],
                }
            )

    def _reply(self, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
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


def serve(port: int = DEFAULT_PORT, host: str = "127.0.0.1"):
    print(f"offipy server listening on http://{host}:{port}", flush=True)
    # 主线程初始化 COM 一次并保持整个服务生命周期
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
    a = ap.parse_args()
    serve(a.port, a.host)
