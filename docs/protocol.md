> [English](protocol.en.md)

# offipy HTTP 协议（P2-8）

offipy 通过本地 HTTP server（默认 `127.0.0.1:8890`）驱动真实 Office。本文定义
`offipy-http/v1` 协议的请求/响应契约、鉴权与版本握手。实现见 `src/offipy/server.py`
（服务端）与 `src/offipy/client.py`（客户端）。

## 版本握手

- 协议名常量：`offipy-http/v1`（`server._PROTOCOL` / `client.PROTOCOL`）。
- 请求侧：`/call` 与 `/shutdown` 必须携带请求头 `X-Offipy-Protocol: offipy-http/v1`。
  - 缺失或值不匹配 → 400，`error_code: "protocol"`（映射回 `ProtocolError`）。
  - 这是**请求侧握手**：旧 client 连新 server（或反超）时协议协商失败，
    不静默错位。
- 服务端：`/status` 响应的 `result.protocol` 报告服务端协议版本，供 client 探测
  （`client._probe` 以此判定 server 是否「我们的」）。

## 端点

| 端点 | 方法 | 鉴权 | 说明 |
|------|------|------|------|
| `/ping` | GET | 否 | 健康检查，返回 `{"ok": true, "result": "pong"}`，不暴露任何数据 |
| `/status` | GET | 是 | 进程/协议/会话标识/目标身份只读快照 |
| `/call` | POST | 是 | 执行一个 RPC 操作（COM op 入 worker 队列串行执行） |
| `/shutdown` | POST | 是 | 鉴权过的优雅停机（回包后由独立线程触发 shutdown，不依赖 pid 强杀） |

其余路径一律 404。

## 鉴权

所有受保护端点需携带 `Authorization: Bearer <token>`。token 来源：环境变量
`OFFIPY_SERVER_TOKEN` 优先，否则 `user_data_dir()/token` 持久文件。校验失败仅 401，
**不杀 server**（token 失配是配置问题，不是进程问题）。

## /call 请求

```json
POST /call
Authorization: Bearer <token>
Content-Type: application/json
X-Offipy-Protocol: offipy-http/v1

{"app": "excel", "op": "set_cell", "args": {"sheet": 1, "cell": "A1", "value": 100}}
```

| 字段 | 说明 |
|------|------|
| `app` | `excel` / `word` / `ppt`（未知 → 400 `invalid_argument`） |
| `op` | 必须是 schema 白名单内 op（`server._OPS` 由 `schema.py` 派生；未知 → 400 `invalid_argument`） |
| `args` | 透传给 App 方法的关键字参数；文件路径类参数（`path`/`out` 等）由 client 按调用方 CWD 绝对化 |

边界（全部在 handler 线程 fail-fast，不触碰 COM）：
- `Content-Type` 必须为 `application/json`，否则 415。
- 请求体上限 16MB，超限 413。
- 负 `Content-Length` → 400（`read(负值)` 会吞掉连接缓冲）。

## /call 响应

统一为 OperationResult 契约（`src/offipy/result.py`）。**这是 HTTP-only 契约**——
Python API / MCP 各有自己的返回形状（Python 返回方法原值、MCP 返回 `data` 载荷），
三入口对照见 [`docs/api.md`](api.md)。

成功（HTTP 200）：

```json
{"ok": true, "operation": "excel.set_cell", "resource_id": "excel:book:book1",
 "message": "ok", "data": null, "result": null}
```

- `resource_id`：`"<app>:<kind>:<doc_id>"` 标识本次操作作用的文档（原始 COM 对象
  不外泄，由 `resource_id` 替代）；**doc_id 是会话内稳定标识，不用用户可改的 name**；
  无目标时为 `null`。
- `data`：操作结果（读 op 的原值；void op 为 `null`）。
- `result`：`data` 的兼容别名（旧 client 渐进切换）。

失败（HTTP 500）：

```json
{"ok": false, "operation": "excel.set_cell", "resource_id": null,
 "error": "TargetNotFoundError: 没有打开的工作簿", "error_code": "target_not_found",
 "trace": ["..."]}
```

- `error_code` 与领域异常一一对应（见 `exceptions.py`）：

| error_code | 领域异常 |
|------------|----------|
| `invalid_argument` | `InvalidArgumentError` |
| `target_not_found` | `TargetNotFoundError` |
| `file_conflict` | `FileConflictError` |
| `com_operation` | `ComOperationError`（保留 `hresult` 字段） |
| `protocol` | `ProtocolError` |
| `internal` | 无映射（未知/普通异常） |

- client 按 `error_code` 把响应映射回对应领域异常，三入口（Python/RPC/MCP）同源。

## 附加字段

- **warning**：op 在 schema 中标 `deprecated` 时，成功/失败响应都带 `warning`
  字段（见 `docs/deprecation.md`）。
- **destructive 确认**：带 `overwrite` 参数的破坏性 op 由 `paths.ensure_writable`
  统一施覆盖保护——目标文件已存在且 `overwrite=false` → `FileConflictError`
  （`error_code: "file_conflict"`）。`expected_target` 用于破坏性 op 的目标绑定：
  **resolve-once**——`{doc_id}` / `{name}` / `{path}` 三键（可组合），用 `get_target(doc_id=...)`
  解析目标，校验 name/path 后把解析出的 doc_id 注入方法调用（杜绝「校验 A 执行 B」）；
  空对象或含未知键直接拒绝 `invalid_argument`，绑定失败 `target_not_found`
  （见 `SECURITY.md`）。

## /status 响应

```json
{"ok": true, "result": {
  "version": "0.9.0a1",
  "protocol": "offipy-http/v1",
  "session_id": "<uuid4>",
  "pid": 28776,
  "python": "3.12.10",
  "started_at": 1785858307.49,
  "targets": {"excel": null, "word": null, "ppt": {"app": "ppt", "doc_id": "pres1", "name": "...", "path": "..."}}
}}
```

- `session_id`：本次 server 会话标识（uuid4），随每条操作日志（`oplog.jsonl`）写入，
  供跨实例区分/追查。
- `targets`：各 App 当前激活文档的只读身份快照，由 worker 线程缓存；handler 不触碰 COM，
  `GET /status` 绝不因探测拉起 Office。

## 操作日志（P2-3）

server 每次 `/call` 后向 `user_data_dir()/oplog.jsonl` 追加一条 JSONL：

```json
{"ts": "...", "session_id": "<uuid4>", "app": "excel", "op": "set_cell",
 "ok": true, "error_code": null, "duration_ms": 12, "resource_id": "excel:book:book1"}
```

args 一律不落盘（脱敏）。日志 ~5MB 轮转（保留 `.1` 备份）。读取：`offipy log` /
`offipy log --tail N`。
