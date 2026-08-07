---
name: Bug report
about: 报告 offipy 的 Bug（direct / Remote / CLI / MCP / deck 转换管线均可）
title: "[Bug] "
labels: bug
---

> 填得越全，修得越快。`环境` 与 `复现步骤` 必填；`故障域自检` 逐项打勾——每勾一项就少一轮排查往返。
> 模版里每一项都对应一个真实踩过的坑，能勾掉说明这条已排除；勾不掉的项请贴对应输出。

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
- **涉及的 op**（如 `excel.set_range` / `deck.render` / `ppt.quit`）：________
  精确到 op 能直接定位代码路径；不确定就写「报错信息里出现的函数名」。
- `offipy check --json` 输出：完整粘贴（脱敏 token）
- `offipy server status` 输出：完整粘贴（脱敏 token；涉及 server / Remote / CLI / MCP 时**必填**——
  它暴露 server 的 version，能直接识别「旧版/stale server」这一常见根因）
- `<app> list_docs` 输出：完整粘贴（涉及 doc_id / 目标文档 / 「找不到文档」类报错时**必填**——
  doc_id 只在同会话内有效；中文 Office 默认文档名是「工作簿N」/「文档N」，不是 BookN/DocN）

## 问题描述

一句话说清现象。

- **出现频率**：必现 / 偶发（约几次出现一次）/ 首次出现。
  （偶发多为竞态 / 残留进程 / 超时；必现多为逻辑错误——两者排查分叉不同，务必如实填写）

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

## 故障域自检（逐项打勾）

> 非必填，但每勾一项姐姐就少追问一轮。勾不掉的项是嫌疑点，请贴对应输出。

### 环境 / 进程

- [ ] 已关闭全部 Office 进程（`tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` 无输出）后问题仍复现
  —— 排除**残留 Office 进程持有文件锁**（save / save_pdf 报「文件被占用」时尤其要查；PowerPoint 加载项常让 quit 后进程残留）
- [ ] 已重启 server（`offipy server stop` 后重试，client 会自动拉起新 server）问题仍复现
  —— 排除**旧版/stale server 加载旧代码**（升级 offipy 或改动源码后最常见；`server status` 的 version 与 `offipy.__version__` 不一致即命中）
- [ ] 出现**连接类错误**（连接失败 / 502 / 超时）时，已确认未开 VPN / 系统代理，或关掉后仍复现
  —— 排除**本地回环被系统代理劫持**（VPN 常劫持 127.0.0.1 回环，表现为连接失败）
- [ ] 中文 Office 环境按名称取文档时，已用 `list_docs()` 核对实际名称
  —— 排除**中文版命名差异**（默认「工作簿N」/「文档N」/「演示文稿N」，不是 BookN/DocN/PresN）

### 入口 / 契约

- [ ] 换一个入口（如 CLI 报错改用 `direct`，或反之）问题同样触发
  —— 排除**入口隔离**（direct 与会话式 server 是两套物理隔离的 doc 注册表，doc_id 互不相通）
- [ ] 报「找不到文档 / doc_id 无效」时，该文档已在**同一次调用链 / 同一会话**内创建或打开
  —— doc_id 只在同会话内有效，跨会话 / 跨入口引用必然失效

### deck 管线

- [ ] 源 HTML 里引用的图片路径**真实存在**
  —— 本地缺失 `<img>` 会静默生成占位图嵌入 PPTX，产物看不出差别（render 产物图片空白/缺失 = 源文件缺失，转换器不报错）
- [ ] `offipy check --json` 的 `chromium` 项为 `ok`
  —— deck 转换依赖 Playwright Chromium（`playwright install chromium`），缺失会在渲染阶段报错

### 回归

- [ ] 升级 offipy 之后才出现（升级前版本：________）
- [ ] 换全新 Office 进程 / 文档问题同样触发（排除僵尸进程 / 损坏文档）

## 补充

- 是否同时跑多个 offipy server（自定义 `--port` / `OFFIPY_SERVER_PORT`）？—— 多实例各自独立 token/pid/oplog，端口冲突会拒双启
- 是否与特定文档/文件相关？能提供最小复现文件吗（脱敏后，注意删除隐私/内部内容）？
- 跑 Python 脚本输出中文乱码？—— 子进程需设 `PYTHONIOENCODING=utf-8`（Windows 默认 GBK），这是环境问题不是 offipy bug
- **安全提醒**：请勿粘贴 `OFFIPY_SERVER_TOKEN` 环境变量或 `user_data_dir()/token` 文件内容等任何密钥。
