> [English](exceptions.en.md)

# 异常契约

库层失败一律抛可捕获的 `OffipyError` 子类（策略 A：完整领域异常），绝不抛 `SystemExit`。
三入口（Python API / RPC / MCP）同源：每个异常带 `code`，server 失败响应带
`error_code`，client 按表映射回对应领域异常。

## 异常表

| 异常 | `code` | 触发场景 |
| --- | --- | --- |
| `OffipyError` | `offipy` | 所有异常基类 |
| `InvalidArgumentError` | `invalid_argument` | 参数/输入非法：单元格解析、常量表、范围校验失败 |
| `TargetNotFoundError` | `target_not_found` | 目标不存在：没有打开的工作簿/文档/演示文稿；未知 `doc_id`；`expected_target` 绑定不匹配 |
| `FileConflictError` | `file_conflict` | 目标文件已存在且未显式 `overwrite=True`（另继承 `FileExistsError`） |
| `ComOperationError` | `com_operation` | App 方法内 COM 调用失败；`hresult` 保留底层 HRESULT 供断连识别 |
| `ProtocolError` | `protocol` | 请求/响应协议版本不匹配，或握手失败 |
| `OfficeUnavailableError` | `office_unavailable` | Office 应用/COM 运行时不可用 |
| `ServerStartError` | `server_start` | 本地常驻 server 无法启动或拉起超时 |
| `RemoteCallError` | `remote_call` | 对常驻 server 的远端调用失败（op 抛错/超时/网络异常） |
| `ConversionError` | `conversion` | HTML→PPTX 转换/渲染失败（含 chromium 缺失） |
| `UnsupportedPlatformError` | `unsupported_platform` | 非 Windows 平台调用 Windows 专属能力 |

## 兼容继承

`InvalidArgumentError` 同时继承 `ValueError`，`FileConflictError` 同时继承
`FileExistsError`：既有 `except ValueError` / `except FileExistsError` 的调用方无需改动。

## 边界处理

- **CLI**：在边界捕获这些异常，转成退出码 + stderr 提示。
- **MCP server**：捕获后转成工具错误返回给模型（`is_error: true`）。
- **库调用方**：直接捕获对应领域异常即可。

## expected_target

破坏性操作（schema 中 `destructive=True`）支持 `expected_target` 目标绑定，键取
`doc_id` / `name` / `path`（可组合）。server 在 dispatch 时 **resolve-once**：用
`get_target(doc_id=...)` 解析出目标 doc_id，校验 `name`/`path` 匹配后，把解析结果
直接注入方法调用参数（杜绝「校验 A 执行 B」）。空对象或含未知键 → `InvalidArgumentError`，
绑定失败（目标不存在 / name/path 不匹配）→ `TargetNotFoundError`。绑定目标是**绑定目标**，
不跟随用户焦点，防止对错误文档执行破坏性操作。
