---
name: Bug report
about: 报告 offipy 的 Bug（COM 操作 / 转换管线 / server / CLI / MCP 均可）
title: "[Bug] "
labels: bug
---

## 环境（必填）

- Windows 版本：`winver` 结果，如 Windows 11 23H2
- Python 版本：`python --version`
- Office 版本与位数：Word/Excel/PowerPoint 的「关于」对话框，如 Microsoft 365 64 位
- offipy 版本：`uv run python -c "import offipy; print(offipy.__version__)"`
- 使用方式：direct / CLI / MCP（三选一，多入口都触发则全部列出）
- `offipy check --json` 输出：请完整粘贴（脱敏 token）

## 问题描述

一句话说清现象。

## 复现步骤

1. …
2. …
3. …

## 期望行为

## 实际行为

## 完整 traceback / 日志

```text
（完整粘贴，不要截断）
```

## 补充

- 是否只在特定 Office 版本/位数出现？
- 是否与特定文档/文件相关？能提供最小复现文件吗（脱敏后）？
