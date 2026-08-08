"""OperationResult：server /call 的返回契约（HTTP-only，P1-3）。

这是 HTTP 层的 /call 响应契约，不是 Python/MCP 的统一返回体——Python API 返回
App 方法原值，MCP 解出 data 载荷后返回（对照见 docs/api.md）。成功返回
{ok, operation, resource_id, message, data}；失败返回 {ok, operation, error,
error_code, ...}。原始 COM 对象不外泄——server 侧 _serialize 把 COM 置 None，
resource_id 替代对象引用，让调用方能定位操作作用于哪个文档/工作簿/演示文稿。

兼容：旧 client 读 result；新响应同时带 data 与 result 别名，渐进切换。
"""

from dataclasses import dataclass
from typing import Any


@dataclass
class OperationResult:
    """一次 /call 操作的结果载体（可序列化契约）。"""

    ok: bool
    operation: str  # "excel.set_cell"
    resource_id: str | None  # "excel:book:book<hex>"（doc_id，会话内稳定标识）
    message: str
    data: Any = None

    def to_dict(self) -> dict:
        d = {
            "ok": self.ok,
            "operation": self.operation,
            "resource_id": self.resource_id,
            "message": self.message,
            "data": self.data,
        }
        if not self.ok:
            d["error"] = self.message
        return d
