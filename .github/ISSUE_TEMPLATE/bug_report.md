---
name: Bug report
about: 报告 offipy 的 Bug（direct / Remote / CLI / MCP / deck 转换管线均可）
title: "[Bug] "
labels: bug
---

> 填得越全，修得越快。`环境` 与 `复现步骤` 必填；`故障域自检` 尽量逐项回答——每多勾一项，
> 就少一轮排查往返。

## 环境（必填）

- Windows 版本：`winver` 结果，如 Windows 11 23H2
- Python 版本：`python --version`
- Office 版本与位数：Word/Excel/PowerPoint 的「关于」对话框，如 Microsoft 365 64 位
- offipy 版本：`python -c "import offipy; print(offipy.__version__)"`（源码开发环境用 `uv run python -c "..."`）
- 安装方式：PyPI（`pip install offipy[all]`）/ 源码开发安装（`uv run`）/ 其他
- **受影响入口**（勾选全部触发的；入口决定调试路径——server 相关 vs 本地 COM）：
  - [ ] `direct`（`offipy.direct.Excel/Word/Ppt`，本地直连 COM，不经 8890 server）
  - [ ] `Remote*`（`offipy.RemoteExcel/RemoteWord/RemotePpt`，经 8890 server）
  - [ ] CLI（`offipy <app> <op> ...`，经 8890 server）
  - [ ] MCP（MCP 工具调用，经 8890 server）
  - [ ] deck 管线（HTML→PPTX：`convert` / `render` / `audit` / `art`）
  - [ ] server 本身（启动 / 状态 / 停机）
- `offipy check --json` 输出：完整粘贴（脱敏 token）
- `offipy server status` 输出：完整粘贴（脱敏 token；涉及 server / Remote / CLI / MCP 时**必填**——
  它暴露 server 的 version，能直接识别「旧版/stale server」这一常见根因）

## 问题描述

一句话说清现象。

## 复现步骤

1. …（最小步骤，能独立跑最好）
2. …
3. …

**最小复现代码 / 命令**（可选但强烈建议）：完整可运行的片段，不要只贴截图。

## 期望行为

## 实际行为

## 完整 traceback / 日志

```text
（完整粘贴，不要截断；保留异常类型、HRESULT（如 0x8004...）、error_code（如 invalid_argument））
```

## 故障域自检（帮助缩小范围，逐项打勾）

> 非必填，但每勾一项姐姐就少追问一轮。

- [ ] 已关闭全部 Office 进程（`tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` 无输出）后问题仍复现
- [ ] 已重启 server（`offipy server stop` 后重试，client 会自动拉起新 server）问题仍复现
  —— 排除**旧版/stale server 加载旧代码**
- [ ] 换一个入口（如 CLI 报错改用 `direct`，或反之）问题同样触发
  —— 排除**入口隔离**（direct 与会话式 server 是两套物理隔离的 doc 注册表，doc_id 互不相通）
- [ ] 换全新 Office 进程 / 文档问题同样触发
  —— 排除**僵尸进程 / 文件锁**（save / save_pdf 报「文件被占用」时尤其要查）
- [ ] 升级 offipy 之后才出现（升级前版本：______）

## 补充

- 是否与特定文档/文件相关？能提供最小复现文件吗（脱敏后，注意删除隐私/内部内容）？
- **安全提醒**：请勿粘贴 `OFFIPY_SERVER_TOKEN` 环境变量或 `user_data_dir()/token` 文件内容等任何密钥。
