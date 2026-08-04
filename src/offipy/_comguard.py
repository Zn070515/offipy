"""COM 错误包装：把 App 公开方法里的 pywintypes.com_error 统一转成 ComOperationError。

dispatch 据此识别断连 HRESULT 触发重建重试；client 也能把 COM 失败映射回
领域异常，而不是让 pywintypes.com_error 裸穿透到三入口。

非 Windows / 无 pywin32 时 _COM_ERROR 降级为空元组 ()——`except ()` 匹配不到
任何异常，guard 原样放行，保证跨平台 import 与纯逻辑单测不炸。
"""

import functools

from .exceptions import ComOperationError

try:
    import pywintypes

    _COM_ERROR = pywintypes.com_error
except ImportError:
    _COM_ERROR = ()


def guard_com(cls):
    """类装饰器：包裹 App 全部公开方法（跳过 `_` 私有与已包装者）。"""
    for name in dir(cls):
        if name.startswith("_"):
            continue
        attr = getattr(cls, name)
        if not callable(attr) or getattr(attr, "_offipy_guarded", False):
            continue
        setattr(cls, name, _guarded(attr))
    return cls


def _guarded(fn):
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except _COM_ERROR as e:
            raise ComOperationError(str(e), hresult=getattr(e, "hresult", None), cause=e) from e

    wrapper._offipy_guarded = True
    return wrapper
