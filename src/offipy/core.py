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
import subprocess
import sys
import time
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, TypeVar

from .exceptions import (
    ComOperationError,
    InvalidArgumentError,
    OfficeUnavailableError,
    TargetNotFoundError,
    UnsupportedPlatformError,
)

_T = TypeVar("_T")

PROGIDS = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "ppt": "PowerPoint.Application",
}

# 各应用 DisplayAlerts「抑制全部弹窗」值：对齐 ppt.PP_ALERTS_NONE(1) /
# word.WD_ALERTS_NONE(0) / Excel False。core 被 app 模块依赖，不能反向
# import 各 app 常量，用字面量 + 注释锁定。Quit 前必须先抑制，否则有未
# 保存文档时 Quit 弹模态保存框 → COM 调用阻塞挂死（H5）。
_ALERTS_NONE = {"ppt": 1, "word": 0, "excel": False}

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


def _target_guard(fn: Callable[..., _T], msg: str) -> Callable[..., _T]:
    """目标绑定守卫工厂（P0-3 doc_id 权威）：强制「显式 doc_id 或 follow_active=True」。

    destructive 与 requires_target 共用同一逻辑：doc_id 缺失且未开 follow_active
    → InvalidArgumentError，绝不静默落到「当前活动文档」（防止用户看到 B、Agent
    改 A）。follow_active 开启时用 self.get_target() 实时解析真实活动目标并注入
    doc_id。差异只在报错文案（msg）：破坏性=改文档，requires_target=写文件/导出。
    """
    sig = inspect.signature(fn)

    @functools.wraps(fn)
    def wrapper(self: Any, *args: Any, follow_active: bool = False, **kw: Any) -> _T:
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
                raise InvalidArgumentError(msg)
        return fn(*bound.args, **bound.kwargs)

    return wrapper


def destructive(fn: Callable[..., _T]) -> Callable[..., _T]:
    """破坏性操作守卫：强制「显式 doc_id 或 follow_active=True」（改源文档）。"""
    return _target_guard(fn, "破坏性操作需要显式 doc_id 或 follow_active=True")


def requires_target(fn: Callable[..., _T]) -> Callable[..., _T]:
    """导出/写文件操作守卫：不修改源文档但写文件系统，同样强制绑定目标。"""
    return _target_guard(fn, "导出/写文件操作需要显式 doc_id 或 follow_active=True")


def readonly_guard(fn: Callable[..., _T]) -> Callable[..., _T]:
    """只读操作守卫：可选 follow_active（#25：只读 op 对齐破坏性语义）。

    只读 op 的 doc_id 缺省本就实时解析当前活动文档；follow_active=True 是显式
    声明该意图，经 get_target() 解析真实活动目标并注入 doc_id，无活动文档抛
    TargetNotFoundError——与 destructive/requires_target 的 follow_active 完全
    一致，让 api.op()/with Excel() 等入口对只读 op 传 follow_active 不再 TypeError。
    """
    sig = inspect.signature(fn)

    @functools.wraps(fn)
    def wrapper(self: Any, *args: Any, follow_active: bool = False, **kw: Any) -> _T:
        bound = sig.bind_partial(self, *args, **kw)
        did = bound.arguments.get("doc_id")
        if did is None and follow_active:
            tgt = self.get_target()
            if tgt is None:
                raise TargetNotFoundError("没有活动文档；请先 new_*/open_* 或显式 doc_id")
            bound.arguments["doc_id"] = tgt["doc_id"]
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

# 0x800401F0：当前线程从未 CoInitialize——是线程契约问题，不是「连不上实例」。
# 单独识别并给 com_apartment() 提示，绝不误归入 _NOT_RUNNING_HRS（否则会被当成
# 「无实例」去 launch，进一步报「无法启动」更误导）。
_CO_E_NOTINITIALIZED = -2147221008


def connect(app: str) -> Any:
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
        # 负 hresult 显示 two's complement（0x800401f0），与微软 HRESULT 文档对照，
        # 而非取绝对值加负号（-0x7ffbfe10 在 32 位下有符号解释为另一个值）。
        fmt = f"0x{hr & 0xFFFFFFFF:08x}" if isinstance(hr, int) else str(hr)
        if hr == _CO_E_NOTINITIALIZED:
            raise ComOperationError(
                f"COM 未初始化: 无法连接 {_progid(app)}（HRESULT {fmt}）——当前线程未调用"
                " CoInitialize。请将调用放入 com_apartment() 套间（P1-4 线程契约）",
                hresult=hr,
            ) from e
        raise ComOperationError(f"无法连接 {_progid(app)}: HRESULT {fmt}", hresult=hr) from e


def launch(app: str, visible: bool = True) -> Any:
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


def ensure_app(
    app: str,
    visible: bool = True,
    modify_existing_visibility: bool = False,
) -> tuple[Any, bool]:
    """优先重连已运行实例，否则新建。返回 (obj, created)。

    连到既有实例时默认不改其可见性（P1-2：用户正在用的窗口不被抢改）；
    modify_existing_visibility=True 才 _set_visible。_set_usercontrol 始终
    保留（移交用户控制，避免实例随 Python 进程退出）。

    COM/Office 运行时不可用时抛 OfficeUnavailableError。
    """
    obj = connect(app)
    if obj is not None:
        if modify_existing_visibility:
            _set_visible(obj, visible)
        _set_usercontrol(obj)
        return obj, False
    try:
        return launch(app, visible), True
    except Exception as e:
        # #88：gen_py 类型库缓存损坏（%TEMP%\gen_py 缺 __init__.py，其中定义
        # CLSIDToClassMap，通常由中断的首次 EnsureDispatch 生成造成）时，裸 AttributeError
        # 只透出 typelib GUID，Agent 无法定位。识别特征并给「删除缓存」修复指引。
        if isinstance(e, AttributeError) and "gen_py" in str(e) and "CLSIDToClassMap" in str(e):
            raise OfficeUnavailableError(
                "pywin32 的 gen_py 类型库缓存已损坏（Office 升级 / 上次类型库生成被"
                "中断 / pywin32 重装常见），EnsureDispatch 无法解析 COM 成员。请删除"
                " %TEMP%\\gen_py 目录后重试（下次启动自动重建类型库）。原始错误: "
                f"{e}"
            ) from e
        raise OfficeUnavailableError(f"无法启动 {_progid(app)}: {e}") from e


# 各应用取主窗口句柄的属性路径：Excel 用 Application.Hwnd；Word/PowerPoint
# 用 ActiveWindow.Hwnd（无活动文档/启动早期拿不到 → None，跳过精确清理）。
_APP_HWND_PATH = {
    "excel": ("Hwnd",),
    "ppt": ("ActiveWindow", "Hwnd"),
    "word": ("ActiveWindow", "Hwnd"),
}


def app_process_pid(obj: Any, app: str) -> int | None:
    """经 COM 窗口句柄反查进程 PID（只用于精确清理本库附着过的实例）。

    拿不到句柄（无活动窗口/断连/启动早期）→ None，调用方跳过清理——
    绝不按进程名模糊清理，避免误杀用户其它 Word/PowerPoint 实例。
    """
    try:
        import ctypes

        cur = obj
        for attr in _APP_HWND_PATH.get(app, ()):
            cur = getattr(cur, attr)
        hwnd = int(cur)
        if not hwnd:
            return None
        getpid = ctypes.windll.user32.GetWindowThreadProcessId
        getpid.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong))
        getpid.restype = ctypes.c_ulong
        pid = ctypes.c_ulong()
        getpid(ctypes.c_void_p(hwnd), ctypes.byref(pid))
    except Exception:
        return None
    else:
        return pid.value or None


def _pid_running(pid: int) -> bool:
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
        return str(pid) in r.stdout
    except Exception:
        return True  # 无法确认 → 假定存活，交由 taskkill 兜底


def wait_process_exit(pid: int | None, timeout: float = 2.0) -> bool:
    """轮询指定进程是否已退出；timeout 内退出 → True，仍存活/无法确认 → False。"""
    if not pid:
        return True
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_running(pid):
            return True
        time.sleep(0.2)
    return not _pid_running(pid)


def reap_process(pid: int | None) -> None:
    """taskkill /F /PID 强制终止指定进程；进程已退/无权限/失败 → 静默。"""
    if not pid:
        return
    with contextlib.suppress(Exception):
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )


def quit_app(app: str) -> bool:
    """退出指定应用的存活实例。无实例时返回 False。

    调用 obj.Quit() 后确认进程退出；Excel 常驻（RCW/COM server 保持）会在
    Quit 返回后残留进程，按 PID 精确清理本实例，避免反复开合累积 Office 进程。
    """
    com = _com()
    obj = connect(app)
    if obj is None:
        return False
    pid = app_process_pid(obj, app)
    with contextlib.suppress(com.pywintypes.com_error):
        obj.DisplayAlerts = _ALERTS_NONE.get(app, False)  # H5：Quit 前抑制弹窗防挂死
        obj.Quit()
    if not wait_process_exit(pid, timeout=2.0):
        reap_process(pid)
    return True


def running(app: str) -> bool:
    """该应用当前是否有存活实例。"""
    return connect(app) is not None


def active_doc(app: str, attr: str) -> Any:
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


def doc_alive(obj: Any) -> bool:
    """缓存文档句柄的 liveness probe：仍连着存活实例才可用。

    COM 不可用（非 Windows / 缺 pywin32）无法探测 → 返回 False，
    调用方（如 quit 的「已退出」判定）把它当 False 即走安全路径。
    """
    try:
        com = _com()
    except UnsupportedPlatformError:
        return False
    try:
        _ = obj.Application.Visible
    except (AttributeError, com.pywintypes.com_error):
        return False
    else:
        return True


def _set_visible(obj: Any, visible: bool) -> None:
    # 个别应用/版本不允许在启动早期设 Visible，静默忽略，
    # 窗口是否可见由 Office 自身决定。
    com = _com()
    with contextlib.suppress(AttributeError, com.pywintypes.com_error):
        obj.Visible = visible


def _set_usercontrol(obj: Any) -> None:
    # UserControl=True 是关键：让 Office 应用独立于 COM 客户端存活，
    # 否则 Python 进程退出时应用会被回收，窗口随之关闭。
    com = _com()
    with contextlib.suppress(AttributeError, com.pywintypes.com_error):
        obj.UserControl = True
