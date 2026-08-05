"""本地直连 COM 的 offipy 入口（P0-4 会话模型：与 Remote* 远程会话相对）。

Excel()/Word()/Ppt() 直接在当前进程内持有 Office COM 句柄，doc_id/线程/
会话状态与 CLI/MCP/Remote* 完全隔离。需要与 CLI 共享同一 Office 会话时，
用 offipy.RemoteExcel/RemoteWord/RemotePpt。
"""

from .api import Excel, Ppt, Word

__all__ = ["Excel", "Word", "Ppt"]
