"""COM 应用生命周期与会话管理。

会话式驱动的核心设计：每次 CLI 调用是独立 Python 进程，通过
GetActiveObject 重连同一个已运行的 Office 实例，实现跨调用保持
窗口可见、持续驱动同一个文档/工作簿/演示文稿。
"""

import contextlib

import pywintypes
import win32com.client
from win32com.client import gencache

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


def _progid(app: str) -> str:
    key = _ALIASES.get(app.lower(), app.lower())
    if key not in PROGIDS:
        raise ValueError(f"不支持的应用: {app}，可选 {list(PROGIDS)}")
    return PROGIDS[key]


def connect(app: str):
    """重连已运行的 Office 实例；没有存活实例时返回 None。"""
    try:
        return win32com.client.GetActiveObject(_progid(app))
    except pywintypes.com_error:
        return None


def launch(app: str, visible: bool = True):
    """新建 Office 实例并设置可见性，移交用户控制使其跨进程存活。

    用 gencache.EnsureDispatch（early binding）：COM 成员编译为稳定类，
    避免 dynamic dispatch 解析属性失败；对 Excel 单实例应用会自动 attach
    已运行的实例。
    """
    obj = gencache.EnsureDispatch(_progid(app))
    _set_visible(obj, visible)
    _set_usercontrol(obj)
    return obj


def ensure_app(app: str, visible: bool = True):
    """优先重连已运行实例，否则新建。返回 (obj, created)。"""
    obj = connect(app)
    if obj is not None:
        _set_visible(obj, visible)
        _set_usercontrol(obj)
        return obj, False
    return launch(app, visible), True


def quit_app(app: str) -> bool:
    """退出指定应用的存活实例。无实例时返回 False。"""
    obj = connect(app)
    if obj is None:
        return False
    try:
        obj.Quit()
        return True
    except pywintypes.com_error:
        return False


def running(app: str) -> bool:
    """该应用当前是否有存活实例。"""
    return connect(app) is not None


def _set_visible(obj, visible: bool) -> None:
    # 个别应用/版本不允许在启动早期设 Visible，静默忽略，
    # 窗口是否可见由 Office 自身决定。
    with contextlib.suppress(AttributeError, pywintypes.com_error):
        obj.Visible = visible


def _set_usercontrol(obj) -> None:
    # UserControl=True 是关键：让 Office 应用独立于 COM 客户端存活，
    # 否则 Python 进程退出时应用会被回收，窗口随之关闭。
    with contextlib.suppress(AttributeError, pywintypes.com_error):
        obj.UserControl = True
