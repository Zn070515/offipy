---
name: Feature request
about: 提议 offipy 的新能力 / 新操作 / 新平台支持
title: "[Feature] "
labels: enhancement
---

> 写明「属于哪个子系统」+「与既有能力的关系」+「验收标准」，能显著减少需求澄清往返。

## 背景 / 动机

想解决什么问题？在什么场景下需要？（越具体越好，有真实使用案例更佳）

## 期望的能力

- 属于哪个子系统（决定实现路径与涉及文件）：
  - [ ] Excel / Word / PowerPoint 原子操作（如 `set_xxx` / `read_xxx`）
  - [ ] CLI / MCP / Remote* 入口
  - [ ] deck 管线（HTML→PPTX：`convert` / `render` / `audit` / `art`）
  - [ ] server（会话 / 鉴权 / 协议 / 幂等）
  - [ ] 其他
- 期望的调用方式：函数签名 / 命令行 / MCP 工具名，能举例最好。

## 与既有能力的关系

- 是否已有近似能力？（如 `list_docs` / `read_range` / `read_slide_summary` 等）缺的是哪一块？
- 兼容性期望：
  - [ ] **纯新增**：不动既有 API/契约 → 可进 MINOR / PATCH，无迁移负担
  - [ ] **允许小幅调整既有行为** → 需评估破坏面，可能升 MINOR + 迁移文档

## 替代方案

目前有没有绕过方式？效果如何？

## 验收标准（达成什么算完成）

- 写可验证的期望，如「当 X 时，Y 应当返回 Z」。

## 是否愿意参与

- [ ] 我可以帮忙实现 / 提供真机验证环境
