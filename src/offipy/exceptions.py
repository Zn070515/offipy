"""offipy 异常体系：库层失败一律抛可捕获的 OffipyError 子类，绝不抛 SystemExit。

CLI（offipy.cli）在边界捕获这些异常并转成退出码 + stderr 提示；
MCP server 捕获后转成工具错误返回给模型；库调用方（高层 API）直接捕获即可。
"""


class OffipyError(Exception):
    """offipy 所有异常的基类。"""


class OfficeUnavailableError(OffipyError):
    """Office 应用/COM 运行时不可用（未安装、无法启动、无存活实例）。"""


class ServerStartError(OffipyError):
    """本地常驻 server 无法启动或拉起超时。"""


class RemoteCallError(OffipyError):
    """对常驻 server 的远端调用失败（op 抛错/超时/网络异常）。"""


class TargetNotFoundError(OffipyError):
    """目标不存在：没有打开的工作簿/文档/演示文稿，或 expected_target 绑定不匹配。"""


class ConversionError(OffipyError):
    """HTML→PPTX 转换/渲染失败（含 chromium 缺失等前置条件不满足）。"""


class UnsupportedPlatformError(OffipyError):
    """非 Windows 平台调用仅支持 Windows 的 Office/COM 能力。"""
