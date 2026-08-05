> [English](deprecation.en.md)

# 弃用政策（P2-9）

offipy 对已过时的 RPC 操作（op）采用声明式弃用：在 schema 标一个标志，
server 自动在响应里加 `warning`，消费者据此渐进切换，不再各自为政。

## 标记方式

`src/offipy/schema.py` 里对应 op 的 `OpSpec` 加 `deprecated=True`：

```python
"old_op": OpSpec(
    description="…",
    destructive=True,
    deprecated=True,  # ← 已弃用
),
```

新增弃用只改 schema 一处，server 响应行为随之生效（server 白名单、CLI、
MCP 注册均派生自 schema，无需三处同步）。

## server 行为

已弃用 op 的**成功与失败响应都带 `warning` 字段**（`server._deprecation_warning`）：

```json
{"ok": true, "operation": "word.old_op", "data": null,
 "warning": "word.old_op 已弃用（deprecated），将在未来版本移除"}
```

未弃用 op 不带该字段。`warning` 是附加信息，不改变 HTTP 状态码与
`error_code`；client/MCP 收到后应提示用户切换替代 op。

## 生命周期

1. **标记弃用**：`deprecated=True`，进入弃用期。替代 op 已在 schema 中可用
   （description 里注明替代品）。
2. **通知周期**：至少保持一个完整 MINOR 版本，让下游经 `warning` 字段
   感知并迁移。此期间 op 功能不变。
3. **移除**：到下一个不兼容版本（按 SemVer，0.x 阶段升 MINOR）删除该 op
   ——App 方法、schema 条目、测试、文档一并清理，server 白名单随之收窄。
   `CHANGELOG.md` 记录移除。

## 消费者约定

- **client**：`call` 返回 `data` 之外，若响应含 `warning` 应透出（例如 CLI
  打印到 stderr）。
- **MCP**：工具元数据可带废弃提示；调用结果含 `warning` 时透传给宿主。
- **CLI**：执行已弃用 op 时打印 `warning`，不阻断调用。

## 当前状态

`schema.py` 目前**没有**任何 op 标记 `deprecated`——本政策是机制预留，
首个弃用 op 出现时在 `CHANGELOG.md` 记录。

实现：`server._deprecation_warning` / `_success_result` / `_error_result`
（`src/offipy/server.py`），一致性由 `tests/test_server_security.py::test_deprecated_op_gets_warning`
覆盖。
