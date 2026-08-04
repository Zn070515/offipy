"""COM 应用生命周期与会话管理。

会话式驱动的核心设计：每次 CLI 调用是独立 Python 进程，通过
GetActiveObject 重连同一个已运行的 Office 实例，实现跨调用保持
窗口可见、持续驱动同一个文档/工作簿/演示文稿。

跨平台：顶层不 import pywin32，COM 依赖全部经 _com() 惰性加载。
非 Windows 或缺少 pywin32 时抛 UnsupportedPlatformError，
保证 `import offipy` 在任何平台都能成功。
"""

import contextlib
import sys
from dataclasses import dataclass
from typing import Any

from .exceptions import OfficeUnavailableError, UnsupportedPlatformError

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


def connect(app: str):
    """重连已运行的 Office 实例；没有存活实例时返回 None。"""
    com = _com()
    try:
        return com.win32com.GetActiveObject(_progid(app))
    except com.pywintypes.com_error:
        return None


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
