# OfficeForClaude 项目说明

Windows-only 的 Office COM 自动化库（office-kit）：会话式驱动 Word/Excel/PowerPoint，外加 HTML→可编辑 PPTX 管线。

## 环境要点（务必遵守）

- Windows + bash（Git Bash）。命令用 bash 语法。
- 用 `uv` 管理 `.venv`。**本机系统 `python` 是 3.10，务必用 `.venv`（3.12）**：`uv run python ...` / `uv run pytest ...` / `uv run office ...`。
- 调用 Python 子进程前 `export PYTHONIOENCODING=utf-8`，否则 Windows 终端中文乱码。
- 禁在长命令尾部挂 `| tail`/`| head`（管道缓冲会丢日志）；长命令用日志重定向 + Read 工具读。
- 本机 VPN 在注册表写系统代理（`ProxyServer=127.0.0.1:12334`、`ProxyOverride` 为空），会把本地 127.0.0.1 回环请求劫持给代理（返回 502）。本地回环必须直连（`client.py` 的 `ProxyHandler({})` 已处理）；出站请求（如 `git clone`）必要时显式 `export https_proxy=http://127.0.0.1:12334`。

## 操作纪律

- **每次测试/验证通过后，主动关掉拉起的 Office 窗口**：`uv run office quit ppt`（/excel/word）。注意 `quit` 只是 `obj.Quit()`，PowerPoint 进程可能残留（加载项/关闭对话框卡住）——quit 后**确认进程已退**：`tasklist | grep -i POWERPNT`，仍在则 `taskkill //F //PID <pid>`。不留窗口占用。
- 改了 server 依赖的模块（如 `ppt.py`、`server.py`）后，必须重启 8890 的 server 进程（`taskkill` 该 PID 后 CLI 自动重建），否则 server 加载旧代码。
- git 提交按小步：指定具体文件、清晰 message。当前主干 `main`。

## 版本号更新规则

- **单一来源**：版本号只写在 `src/offipy/__init__.py` 的 `__version__`；`pyproject.toml` 经 `[tool.hatch.version]` 自动读取，**别处不重复写**。
- **语义化版本（SemVer）** `MAJOR.MINOR.PATCH`：
  - `MAJOR`：不兼容的 API 变更（改包名、改调用签名、删既有操作）
  - `MINOR`：向后兼容的新功能（新增原子操作 / CLI 子命令 / 管线特性）
  - `PATCH`：bug 修复、内部重构、不改变对外行为的小调整
- **0.x 开发期约定**：破坏性变更升 `MINOR` 即可，不升 `MAJOR`（尚未对外承诺稳定性）。
- **每次升版本**：单独成一个 commit（message 形如 `chore: bump version to 0.2.0`）；若已发布 PyPI，同步打 tag `v0.2.0` 并推送。功能 commit 不带版本号，升版本也不夹带功能改动。
- **验证**：改完跑 `uv run python -c "import offipy; print(offipy.__version__)"` 确认生效。
