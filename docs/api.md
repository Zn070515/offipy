> [English](api.en.md)

# 返回契约与三入口对照

offipy 同一批操作经三条入口暴露：**Python API**（`offipy.api` 的 `Excel()/Word()/Ppt()`
及底层 App 类）、**HTTP RPC**（`/call`，见 `src/offipy/server.py`）、**MCP**（`offipy mcp`，
见 `src/offipy/mcp_server.py`）。三者同源（同一批 schema / App 方法、同一套领域异常），
但**返回形状不同**——本文如实描述，不强行统一成同一种返回体。

## 各入口返回什么

| 入口 | 成功返回 | void op | 失败 |
|------|----------|---------|------|
| **Python API** | App 方法原始返回值（`new_book`→`doc_id` 字符串、`get_cell`→单元值、`get_target`→dict …） | `None` | 抛 `OffipyError` 领域异常（`exceptions.py`） |
| **HTTP RPC `/call`** | `OperationResult` dict `{ok, operation, resource_id, message, data}`（另附 `result` 兼容别名） | 同上，`data: null` | HTTP 状态码按 `error_code` 映射（400/404/409/500/502/503，见 `docs/protocol.md`），body `{ok:false, error, error_code, trace, ...}` |
| **MCP 工具** | 操作 `data` 载荷（读 op 的原值） | `"ok (<op>)"` 字符串 | MCP 错误（与领域异常同源文案） |

示例：`excel.get_cell(1, "A1")`

| 入口 | 写法 | 返回 |
|------|------|------|
| Python | `x.get_cell(1, "A1")` | `100` |
| HTTP | `POST /call {"app":"excel","op":"get_cell","args":{"sheet":1,"cell":"A1"}}` | `{"ok":true, "operation":"excel.get_cell", "resource_id":"excel:book:book2f9a5c5e1a2b3c4d", "message":"ok", "data":100, "result":100}` |
| MCP | `excel_get_cell(sheet=1, cell="A1")` | `100` |

## OperationResult（HTTP-only 契约）

`OperationResult` 是 **server 的 `/call` 响应契约**，定义在 `src/offipy/result.py`。
它不是 Python/MCP 的统一返回体——Python API 直接返回方法原值，MCP 解出 `data` 后返回。

成功（HTTP 200）：

```json
{"ok": true, "operation": "excel.set_cell", "resource_id": "excel:book:book2f9a5c5e1a2b3c4d",
 "message": "ok", "data": null, "result": null}
```

| 字段 | 说明 |
|------|------|
| `ok` | 布尔，操作是否成功 |
| `operation` | `"<app>.<op>"` 式全名，如 `excel.set_cell` |
| `resource_id` | `"<app>:<kind>:<doc_id>"`——标识本次操作作用的文档；**doc_id 是会话内稳定标识，不用用户可改的 name**（`book<hex>`/`doc<hex>`/`pres<hex>` 高熵不透明，不可枚举）；无目标为 `null` |
| `message` | 人类可读信息（成功一般 `"ok"`） |
| `data` | 操作结果：读 op 的原值，void op 为 `null` |
| `result` | `data` 的兼容别名（旧 client 渐进切换） |

失败（HTTP 状态码按 `error_code` 映射，见 `docs/protocol.md` 的表）：

```json
{"ok": false, "operation": "excel.get_cell", "resource_id": null,
 "error": "TargetNotFoundError: 没有打开的工作簿", "error_code": "target_not_found",
 "trace": ["TargetNotFoundError: 没有打开的工作簿"]}
```

`error_code` 与领域异常一一对应（见 `docs/exceptions.md`），client 据此把响应映射回
对应领域异常，三条入口同源——**不管 HTTP 状态码多少，语义一律由 body 的 `error_code`
保证**。`ComOperationError` 额外带 `hresult` 字段；弃用 op 额外带 `warning`
（见 `docs/deprecation.md`）。`trace` 为异常链消息脱敏列表，不含路径/行号/源码。

## 会话语义（目标身份）

op 缺省作用在**当前活动文档**上（ActiveWorkbook / ActiveDocument / ActivePresentation）：

- `get_target`：查询当前激活目标身份 `{"app", "doc_id", "name", "path"}`（无 → `null`）。
- `activate(doc_id)`：把指定文档设为活动目标，并同步真实 UI（Excel `Workbook.Activate()`、
  Word `Document.Activate()`、PPT 激活含该文档的窗口）；同步失败回滚并抛 `ComOperationError`。
- `list_docs`：如实返回已登记句柄的文档表 `{doc_id: {"name", "path", "active"}}`，
  **不隐式枚举**——只报我们持有句柄的文档。
- 破坏性 op 可带 `expected_target`（`{doc_id}` / `{name}` / `{path}` / 组合）做**目标绑定**：
  resolve-once——校验用解析出的 doc_id 注入方法调用，杜绝「校验 A 执行 B」；空对象或含
  未知键直接拒绝（`InvalidArgumentError`），绑定失败抛 `TargetNotFoundError`。

## 安全模型速览

server 只监听 `127.0.0.1`，Bearer token 鉴权（见 `SECURITY.md`）；`server stop` 走鉴权
`/shutdown` 优雅停机，token 失配 / 端口归属无法证明时**一律不杀**，只提示手动处理。
