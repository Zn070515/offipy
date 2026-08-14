"""offipy 异常体系：库层失败一律抛可捕获的 OffipyError 子类，绝不抛 SystemExit。

策略 A（完整领域异常，round-3）：
- 参数/输入非法（解析、常量表、范围校验失败）→ InvalidArgumentError
- 目标不存在 / expected_target 绑定失败 → TargetNotFoundError
- 目标文件已存在且未显式 overwrite → FileConflictError
- App 方法内 COM 调用失败 → ComOperationError（保留 hresult 供断连识别）
- 协议不符 / 握手失败 → ProtocolError

每个异常带 `code`（RPC error_code 与异常一一对应）：server 失败响应带
error_code，client 按表映射回对应领域异常，三入口（Python/RPC/MCP）同源。

CLI（offipy.cli）在边界捕获这些异常并转成退出码 + stderr 提示；
MCP server 捕获后转成工具错误返回给模型；库调用方（高层 API）直接捕获即可。
"""


class OffipyError(Exception):
    """offipy 所有异常的基类。"""

    code = "offipy"


class OfficeUnavailableError(OffipyError):
    """Office 应用/COM 运行时不可用（未安装、无法启动、无存活实例）。"""

    code = "office_unavailable"


class ServerStartError(OffipyError):
    """本地常驻 server 无法启动或拉起超时。"""

    code = "server_start"


class RemoteCallError(OffipyError):
    """对常驻 server 的远端调用失败（op 抛错/超时/网络异常）。"""

    code = "remote_call"


class TargetNotFoundError(OffipyError):
    """目标不存在：没有打开的工作簿/文档/演示文稿，或 expected_target 绑定不匹配。"""

    code = "target_not_found"


class ConversionError(OffipyError):
    """HTML→PPTX 转换/渲染失败（含 chromium 缺失等前置条件不满足）。"""

    code = "conversion"


class UnsupportedPlatformError(OffipyError):
    """非 Windows 平台调用仅支持 Windows 的 Office/COM 能力。"""

    code = "unsupported_platform"


class InvalidArgumentError(OffipyError, ValueError):
    """参数/输入非法（解析、常量表、范围校验失败）。

    同时继承 ValueError：既有 `except ValueError` 的调用方（含历史测试）不变。
    """

    code = "invalid_argument"


class FileConflictError(OffipyError, FileExistsError):
    """目标文件已存在且未显式 overwrite。

    同时继承 FileExistsError：`except FileExistsError` 的既有调用方不变。
    """

    code = "file_conflict"


class ComOperationError(OffipyError):
    """App 方法内 COM 调用失败；hresult 保留底层 HRESULT，供断连识别/重试。"""

    code = "com_operation"

    def __init__(
        self, message: str, *, hresult: int | None = None, cause: BaseException | None = None
    ):
        super().__init__(message)
        self.hresult = hresult
        self.cause = cause


class ProtocolError(OffipyError):
    """协议不符：请求侧握手 / 响应协议版本不匹配（P2-8 落地后由请求侧校验触发）。"""

    code = "protocol"
