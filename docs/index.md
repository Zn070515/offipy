> [English](index.en.md)

# offipy

**offipy** 是一个 Windows-only 的 Office COM 自动化库：会话式驱动
Word / Excel / PowerPoint，外加 HTML→可编辑 PPTX 的转换管线。它把 Office 应用当成
一个「会话」，跨进程复用同一实例，并提供 server / CLI / MCP 三个一致的入口。

## 特性

- **会话式**：跨进程复用同一 Office 实例，`ActiveWorkbook` / `ActiveDocument` /
  `ActivePresentation` 实时定位目标。
- **多文档（P2-2）**：每个应用维护文档表，`doc_id` 显式路由，`activate` / `list_docs`
  管理活动目标；多 server 实例按端口隔离。
- **三入口一致**：HTTP server（127.0.0.1:8890，Bearer token）、CLI（`offipy <app> <op>`）、
  MCP 工具，全部从 `schema.py` 单一来源派生。
- **领域异常**：`InvalidArgumentError` / `TargetNotFoundError` / `FileConflictError` /
  `ComOperationError` / `ProtocolError` 统一继承 `OffipyError`，RPC 带 `error_code`。
- **HTML→PPTX 管线**：Chromium 渲染 HTML 布局 → 可编辑 PPTX，支持图表、图标、主题。

## 安装

```bash
py -m pip install "offipy[all]"     # office + deck + mcp
py -m playwright install chromium   # deck 管线需要
```

按用途拆分：`offipy[office]`（仅 COM 自动化）、`offipy[deck]`（HTML→PPTX）、
`offipy[mcp]`（MCP server）。

## 快速开始

```bash
offipy excel new_book            # 返回 "book1"
offipy excel set_cell --sheet 1 --cell A1 --value 42 --follow-active
offipy excel read_range --sheet 1 --range_addr A1:A1   # [[42.0]]
offipy excel quit
```

MCP 接入 Claude Desktop：参考 README 的 MCP 配置段，指向 `offipy mcp`。

## 文档导航

- [快速上手](usage.md)：会话、多文档、CLI / server / MCP 的完整用法。
- [API 参考](api/index.md)：全部操作（由 schema 自动生成）。
- [异常契约](exceptions.md)：策略 A 领域异常与 RPC `error_code` 映射。
- [协议](protocol.md)：HTTP 协议、token、`/shutdown`、协议版本握手。
- [兼容矩阵](compatibility.md)：Windows / Office / Python / extras 支持情况。
- [弃用政策](deprecation.md)：弃用流程与响应 `warning` 字段。
- [迁移指南](migration.md)：0.9 → 0.10 破坏性变更与迁移步骤。

## 构建文档

```bash
uv run --with mkdocs-material mkdocs build
```

`docs/api/` 由 `scripts/gen_api_ref.py` 从 `schema.py` 自动生成，修改操作后重跑即可。
