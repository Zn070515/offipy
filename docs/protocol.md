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
- 版本偏斜（#34）：协议名匹配但 `result.version` 与 client 的 `__version__` 不一致时，
  `_probe` 判定为 `mismatch`——旧版 server 视为 stale，`ensure_server` 会按 pid 归属
  重启它；`server_status()` 此时仍返回含 `version` 的可读 dict（非 offipy 进程协议失配
  才返回 None）。

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

{"app": "excel", "op": "set_cell",
 "args": {"sheet": 1, "cell": "A1", "value": 100, "doc_id": "book<hex>"},
 "request_id": "2f9a5c5e-0000-4000-8000-000000000001"}
```

| 字段 | 说明 |
|------|------|
| `app` | `excel` / `word` / `ppt`（未知 → 400 `invalid_argument`） |
| `op` | 必须是 schema 白名单内 op（`server._OPS` 由 `schema.py` 派生；未知 → 400 `invalid_argument`） |
| `args` | 透传给 App 方法的关键字参数（含 `doc_id` 目标）；文件路径类参数（`path`/`out` 等）由 client 按调用方 CWD 绝对化。破坏性 op 还可带传输层参数 `expected_target` / `follow_active`（见下） |
| `request_id` | 可选；调用方持有的幂等标识（uuid4 字符串）。提供时 server 按 request_id + payload hash 去重/合并/回放缓存（见「幂等」）；缺省则走不带幂等的老路径 |

边界（全部在 handler 线程 fail-fast，不触碰 COM）：
- `Content-Type` 必须为 `application/json`，否则 415。
- 请求体上限 16MB，超限 413。
- 负 `Content-Length` → 400（`read(负值)` 会吞掉连接缓冲）。
- `args` 非对象（list/str）→ 400 `invalid_argument`。

### 传输层参数（目标绑定）

`expected_target` / `follow_active` 是传输层参数：client 直接透传、不进 App 方法签名，
由 server dispatch 弹出后解析并注入 `doc_id`。`expected_target` 只对破坏性/导出 op
（`schema.supports_expected_target`）有意义；`follow_active` 对破坏性 op 与声明了
`accepts_follow_active` 的只读 op（get_cell / read_range / read_doc_text / read_slide_texts /
read_slide_summary）都有意义（#25：只读 op 对齐破坏性语义，写后读验证场景可用）：

- `follow_active`（bool，可选，默认 `false`）：显式声明「跟随当前活动文档」——server 实时解析
  当前激活目标并注入其 doc_id；无活动目标 → `TargetNotFoundError`（绝不静默落到任何文档）。
- `expected_target`（对象，可选）：目标绑定——`{"doc_id"}` / `{"name"}` / `{"path"}` 可组合，
  **resolve-once**：server 用绑定键解析出目标 doc_id，校验后注入方法调用；空对象 / 含未知键 →
  400 `invalid_argument`；绑定失败 → `target_not_found`。

优先级：`expected_target` > `follow_active` > 显式 `doc_id`（前两者会覆盖 args 里已有的 doc_id）。
正常用法三者取一即可。约束：

- 非破坏性 op 出现 `expected_target` → 400 `invalid_argument`（严格拒绝，不静默忽略）。
- 未声明 `accepts_follow_active` 的只读 op 上的 `follow_active` 静默忽略（被 pop 掉）；
  声明了的只读 op 实时解析活动文档并注入 doc_id（无活动目标 → `TargetNotFoundError`）。
- `quit` 不接受两者（无 doc_id 目标）。

### 幂等（request_id，P0-2 方案 A）

提供 `request_id` 时，server 开启幂等路径——「超时重试不重执行」：

- **payload hash 绑定**：`sha256(json.dumps({"app","op","args"}, sort_keys=True))`。同 request_id
  换了 payload（参数漂移）→ 400 `invalid_argument`（调用方 bug，不静默返回旧结果）。
- **in-flight 合并**：并发/重试带同 request_id 时，非 owner 线程等待 owner 完成（`entry.event.wait`），
  不重复入队、不重复执行。
- **结果缓存**：owner 完成后结果缓存（LRU 上限 512，TTL 600s，与超时窗口同量级），同 request_id
  重试直接回放缓存响应并标注 `cached: true`；`done` 条目被淘汰只发生在非 inflight 时。
- **超时**：owner 等待超 `_CALL_TIMEOUT` → 504，但 entry 留 inflight——同 id 重试仍合并、绝不双写。
  COM 队列满 → 503 并回滚 entry（同 id 重试重建，不 merge 到永不完成）。
- 不带 request_id 的调用走老路径：不入缓存、不去重、不合并。

client 侧：`client.request/call` 缺省自动生成 uuid4 并随请求带上；响应回显 request_id 供调用方核对。
超时重试务必复用同一 request_id。

## /call 响应

统一为 OperationResult 契约（`src/offipy/result.py`）。**这是 HTTP-only 契约**——
Python API / MCP 各有自己的返回形状（Python 返回方法原值、MCP 返回 `data` 载荷），
三入口对照见 [`docs/api.md`](api.md)。

成功（HTTP 200）：

```json
{"ok": true, "operation": "excel.set_cell", "resource_id": "excel:book:book2f9a5c5e1a2b3c4d",
 "message": "ok", "data": null, "result": null,
 "request_id": "2f9a5c5e-0000-4000-8000-000000000001"}
```

- `resource_id`：`"<app>:<kind>:<doc_id>"` 标识本次操作作用的文档（原始 COM 对象
  不外泄，由 `resource_id` 替代）；**doc_id 是会话内稳定标识，不用用户可改的 name**；
  格式 `book<hex>` / `doc<hex>` / `pres<hex>`——**高熵不透明**（`secrets.token_hex(8)`），
  不可顺序枚举，无目标时为 `null`。
- `data`：操作结果（读 op 的原值；void op 为 `null`）。
- `result`：`data` 的兼容别名（旧 client 渐进切换）。
- `request_id`：幂等回显——请求带了 request_id 时原样带回，供调用方核对/重试。

带 request_id 的幂等调用命中缓存时，响应额外带 `"cached": true`（同 request_id 同 payload
重试不重执行，回放原响应）。

失败（HTTP 状态码按 `error_code` 映射，见下表；未列出的 code 回落 500）：

```json
{"ok": false, "operation": "excel.set_cell", "resource_id": null,
 "error": "TargetNotFoundError: 没有打开的工作簿", "error_code": "target_not_found",
 "trace": ["TargetNotFoundError: 没有打开的工作簿"],
 "request_id": "2f9a5c5e-0000-4000-8000-000000000001"}
```

- `error_code` 与领域异常一一对应（见 `exceptions.py`）：

| error_code | 领域异常 | HTTP 状态码 |
|------------|----------|-------------|
| `invalid_argument` | `InvalidArgumentError` | 400 |
| `protocol` | `ProtocolError` | 400 |
| `target_not_found` | `TargetNotFoundError` | 404 |
| `file_conflict` | `FileConflictError` | 409 |
| `com_operation` | `ComOperationError`（保留 `hresult` 字段） | 502 |
| `internal` | 无映射（未知/普通异常） | 500 |

- client 按 `error_code` 把响应映射回对应领域异常，三入口（Python/RPC/MCP）同源——
  不管 HTTP 状态码是多少，语义一律由 body 的 `error_code` 保证；状态码只供监控/
  代理层按语义区分失败类别。
- `trace`：异常链消息脱敏列表（`["类型: 消息", ...]`），只用于定位异常类型与链路；
  不含文件路径/行号/源码片段（服务器信息泄露防护）。
- `error` / `trace` 的**消息内容**同样脱敏（#67）：异常消息里的绝对路径
  （Windows/POSIX/UNC）与 `doc_id` 值统一替换为 `[REDACTED]`，路径原样不再透传。

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
  "version": "0.10.2",
  "protocol": "offipy-http/v1",
  "session_id": "<uuid4>",
  "pid": 28776,
  "python": "3.12.10",
  "started_at": 1785858307.49,
  "targets": {"excel": null, "word": null, "ppt": {"app": "ppt", "doc_id": "pres2f9a5c5e1a2b3c4d", "name": "...", "path": "..."}}
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
 "ok": true, "error_code": null, "duration_ms": 12, "resource_id": "excel:book:book2f9a5c5e1a2b3c4d"}
```

args 一律不落盘（脱敏）。日志 ~5MB 轮转（保留 `.1` 备份）。读取：`offipy log` /
`offipy log --tail N`。
