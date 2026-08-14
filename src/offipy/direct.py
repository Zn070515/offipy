"""本地直连 COM 的 offipy 入口（P0-4 会话模型：与 Remote* 远程会话相对）。

Excel()/Word()/Ppt() 直接在当前进程内持有 Office COM 句柄，doc_id/线程/
会话状态与 CLI/MCP/Remote* 完全隔离。需要与 CLI 共享同一 Office 会话时，
用 offipy.RemoteExcel/RemoteWord/RemotePpt。

线程契约（P1-4）：COM 对象绑定创建它的线程（STA）。direct facade 不是
线程安全的——同一 facade 实例必须在创建它的线程内使用。多线程各自创建
facade 时，每个线程都应包一层 com_apartment()（线程各自 CoInitialize），
让线程级 COM 初始化/反初始化正确成对。
"""

import contextlib
from collections.abc import Iterator

from .api import Excel, Ppt, Word

__all__ = ["Excel", "Ppt", "Word", "com_apartment"]


@contextlib.contextmanager
def com_apartment() -> Iterator[None]:
    """STA COM 套间：direct facade 须在创建它的同一线程内使用。

    跨线程需各自包一层 com_apartment()（线程各自 CoInitialize/CoUninitialize
    成对），避免 COM 对象跨线程访问失败。非 Windows / 无 pythoncom 时是
    no-op（exit 正常）。
    """
    try:
        import pythoncom
    except ImportError:
        yield
        return
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()
