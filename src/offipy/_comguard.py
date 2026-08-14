"""COM 错误包装：把 App 公开方法里的 pywintypes.com_error 统一转成 ComOperationError。

dispatch 据此识别断连 HRESULT 触发重建重试；client 也能把 COM 失败映射回
领域异常，而不是让 pywintypes.com_error 裸穿透到三入口。

非 Windows / 无 pywin32 时 _COM_ERROR 降级为空元组 ()——`except ()` 匹配不到
任何异常，guard 原样放行，保证跨平台 import 与纯逻辑单测不炸。
"""

import functools
import time
from collections.abc import Callable
from typing import Any, TypeVar, cast

from .exceptions import ComOperationError

try:
    import pywintypes

    _COM_ERROR = pywintypes.com_error
except ImportError:
    _COM_ERROR = ()


def guard_com(cls: type[Any]) -> type[Any]:
    """类装饰器：包裹 App 全部公开方法（跳过 `_` 私有与已包装者）。"""
    for name in dir(cls):
        if name.startswith("_"):
            continue
        attr = getattr(cls, name)
        if not callable(attr) or getattr(attr, "_offipy_guarded", False):
            continue
        setattr(cls, name, _guarded(attr))
    return cls


_F = TypeVar("_F", bound=Callable[..., Any])


def _guarded(fn: _F) -> _F:
    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except _COM_ERROR as e:
            raise ComOperationError(str(e), hresult=getattr(e, "hresult", None), cause=e) from e

    # 标记已包裹：wrapper 上挂动态属性，mypy 不认 Callable 的任意属性，用 setattr 绕开
    setattr(wrapper, "_offipy_guarded", True)  # noqa: B010 — setattr 是 mypy 绕开手段，非普通属性赋值
    return cast("_F", wrapper)


# #37：SaveAs/ExportAsFixedFormat 的锁感知短重试。目标文件可能被其他进程（残留的
# Office 实例）持有文件锁，首次保存抛 COM 错误；短重试（200ms×5）后仍失败 → 抛
# 带可读提示的 ComOperationError，而非透传裸 hresult。save_call 为无参 callable
# （调用方已绑定目标/参数），what 用于错误文案（如「保存演示文稿」）。
SAVE_RETRY_ATTEMPTS = 5
SAVE_RETRY_DELAY = 0.2


def _save_attempt(save_call: Callable[[], None]) -> Exception | None:
    """单次保存尝试：成功返回 None；_COM_ERROR 睡短延迟后返回异常供重试。"""
    try:
        save_call()
    except _COM_ERROR as e:
        time.sleep(SAVE_RETRY_DELAY)
        return cast("Exception", e)
    else:
        return None


def save_with_lock_retry(save_call: Callable[[], None], *, what: str) -> None:
    """执行带锁感知短重试的保存调用；超时给可读错误，绝不透传裸 COM 失败。"""
    last: Exception | None = None
    for _ in range(SAVE_RETRY_ATTEMPTS):
        err = _save_attempt(save_call)
        if err is None:
            return
        last = err
    raise ComOperationError(
        f"{what}失败: 目标文件可能被其他进程占用（如残留的 Office 进程持有文件锁）。"
        f"请关闭占用该文件的 Office 进程后重试。原因: {last}"
    ) from last
