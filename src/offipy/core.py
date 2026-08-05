"""COM 应用生命周期与会话管理。

会话式驱动的核心设计：每次 CLI 调用是独立 Python 进程，通过
GetActiveObject 重连同一个已运行的 Office 实例，实现跨调用保持
窗口可见、持续驱动同一个文档/工作簿/演示文稿。

跨平台：顶层不 import pywin32，COM 依赖全部经 _com() 惰性加载。
非 Windows 或缺少 pywin32 时抛 UnsupportedPlatformError，
保证 `import offipy` 在任何平台都能成功。
"""

import contextlib
import functools
import inspect
import sys
from dataclasses import dataclass
from typing import Any

from .exceptions import (
    ComOperationError,
    InvalidArgumentError,
    OfficeUnavailableError,
    TargetNotFoundError,
    UnsupportedPlatformError,
)

PROGIDS = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "ppt": "PowerPoint.Application",
}

_ALIASES = {
    "word": "word",
    "excel": "excel",
    "ppt": "ppt",
    "powerpoint": "ppt",
}


@dataclass(frozen=True)
class _ComBundle:
    """惰性加载的 pywin32 模块束；COM 模块是动态类型，字段标 Any。"""

    pywintypes: Any
    win32com: Any  # win32com.client
    gencache: Any


_COM: _ComBundle | None = None


def _com() -> _ComBundle:
    """惰性加载 COM 依赖；非 Windows 或缺少 pywin32 抛 UnsupportedPlatformError。"""
    global _COM
    if _COM is not None:
        return _COM
    if sys.platform != "win32":
        raise UnsupportedPlatformError("Office COM 自动化仅支持 Windows")
    try:
        import pywintypes
        import win32com.client
        from win32com.client import gencache
    except ImportError as e:
        raise UnsupportedPlatformError(
            "缺少 pywin32，Office COM 自动化不可用（安装: uv pip install pywin32）"
        ) from e
    _COM = _ComBundle(pywintypes, win32com.client, gencache)
    return _COM


def _progid(app: str) -> str:
    key = _ALIASES.get(app.lower(), app.lower())
    if key not in PROGIDS:
        raise ValueError(f"不支持的应用: {app}，可选 {list(PROGIDS)}")
    return PROGIDS[key]


def destructive(fn):
    """破坏性操作守卫（P0-3 doc_id 权威）：强制「显式 doc_id 或 follow_active=True」。

    拦截破坏性 App 方法：doc_id 缺失且未开 follow_active → InvalidArgumentError，
    绝不静默落到「当前活动文档」（防止用户看到 B、Agent 改 A）。follow_active
    开启时用 self.get_target() 实时解析真实活动目标并注入 doc_id。
    """
    sig = inspect.signature(fn)

    @functools.wraps(fn)
    def wrapper(self, *args, follow_active=False, **kw):
        bound = sig.bind_partial(self, *args, **kw)
        did = bound.arguments.get("doc_id")
        if did is None:
            if follow_active:
                tgt = self.get_target()
                if tgt is None:
                    raise TargetNotFoundError("没有活动文档；请先 new_book/open_book 或显式 doc_id")
                did = tgt["doc_id"]
                bound.arguments["doc_id"] = did
            else:
                raise InvalidArgumentError("破坏性操作需要显式 doc_id 或 follow_active=True")
        return fn(*bound.args, **bound.kwargs)

    return wrapper


# GetActiveObject 连不到已运行实例的 HRESULT（有符号 int，与 com_error.hresult 一致）：
# 仅这三类视为「没有可连实例」→ 返回 None 走 launch；其余（权限/RPC/注册损坏）抛
# ComOperationError，绝不静默拉起——否则权限问题会被掩盖成新建实例。
_NOT_RUNNING_HRS = {
    -2147221021,  # 0x800401E3 MK_E_UNAVAILABLE：ROT 里没有该 ProgID 的存活对象
    -2147221164,  # 0x80040154 REGDB_E_CLASSNOTREG：类未注册（Office 未安装）
    -2147221005,  # 0x800401F3 CO_E_CLASSSTRING：类字符串无法映射到 CLSID（Office 未安装）
}


def connect(app: str):
    """重连已运行的 Office 实例；没有存活实例时返回 None。

    仅「未运行 / 类未注册」两类 HRESULT 返回 None（触发 launch）；其余 COM
    失败（权限、RPC 断开、注册表损坏等）抛 ComOperationError，不静默拉起。
    """
    com = _com()
    try:
        return com.win32com.GetActiveObject(_progid(app))
    except com.pywintypes.com_error as e:
        hr = getattr(e, "hresult", None)
        if hr is None and e.args and isinstance(e.args[0], int):
            hr = e.args[0]
        if hr in _NOT_RUNNING_HRS:
            return None
        fmt = f"{hr:#010x}" if isinstance(hr, int) else str(hr)
        raise ComOperationError(f"无法连接 {_progid(app)}: HRESULT {fmt}", hresult=hr) from e


def launch(app: str, visible: bool = True):
    """新建 Office 实例并设置可见性，移交用户控制使其跨进程存活。

    用 gencache.EnsureDispatch（early binding）：COM 成员编译为稳定类，
    避免 dynamic dispatch 解析属性失败；对 Excel 单实例应用会自动 attach
    已运行的实例。
    """
    com = _com()
    obj = com.gencache.EnsureDispatch(_progid(app))
    _set_visible(obj, visible)
    _set_usercontrol(obj)
    return obj


def ensure_app(app: str, visible: bool = True):
    """优先重连已运行实例，否则新建。返回 (obj, created)。

    COM/Office 运行时不可用时抛 OfficeUnavailableError。
    """
    obj = connect(app)
    if obj is not None:
        _set_visible(obj, visible)
        _set_usercontrol(obj)
        return obj, False
    try:
        return launch(app, visible), True
    except Exception as e:  # noqa: BLE001 — win32 异常种类繁多，收拢到语义化异常
        raise OfficeUnavailableError(f"无法启动 {_progid(app)}: {e}") from e


def quit_app(app: str) -> bool:
    """退出指定应用的存活实例。无实例时返回 False。"""
    com = _com()
    obj = connect(app)
    if obj is None:
        return False
    try:
        obj.Quit()
        return True
    except com.pywintypes.com_error:
        return False


def running(app: str) -> bool:
    """该应用当前是否有存活实例。"""
    return connect(app) is not None


def active_doc(app: str, attr: str):
    """会话语义（P1.2）：GetActiveObject 取实时活动文档。

    attr 为 ActivePresentation / ActiveWorkbook / ActiveDocument 之一。
    无运行实例、或该应用当前没有活动文档时返回 None（异常收拢为 None，
    交由调用方决定回退到缓存句柄）。
    """
    com = _com()
    try:
        obj = com.win32com.GetActiveObject(_progid(app))
        return getattr(obj, attr)
    except com.pywintypes.com_error:
        return None


def doc_alive(obj) -> bool:
    """缓存文档句柄的 liveness probe：仍连着存活实例才可用。"""
    com = _com()
    try:
        _ = obj.Application.Visible
        return True
    except (AttributeError, com.pywintypes.com_error):
        return False


def _set_visible(obj, visible: bool) -> None:
    # 个别应用/版本不允许在启动早期设 Visible，静默忽略，
    # 窗口是否可见由 Office 自身决定。
    com = _com()
    with contextlib.suppress(AttributeError, com.pywintypes.com_error):
        obj.Visible = visible


def _set_usercontrol(obj) -> None:
    # UserControl=True 是关键：让 Office 应用独立于 COM 客户端存活，
    # 否则 Python 进程退出时应用会被回收，窗口随之关闭。
    com = _com()
    with contextlib.suppress(AttributeError, com.pywintypes.com_error):
        obj.UserControl = True
