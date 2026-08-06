# offipy 项目开发说明

Windows-only 的 Office COM 自动化库（offipy）：会话式驱动 Word/Excel/PowerPoint，外加 HTML→可编辑 PPTX 管线。

## 环境要点（务必遵守）

- Windows + bash（Git Bash）。命令用 bash 语法。
- 用 `uv` 管理 `.venv`。**本机系统 `python` 是 3.10，务必用 `.venv`（3.12）**：`uv run python ...` / `uv run pytest ...` / `uv run offipy ...`。
- 调用 Python 子进程前 `export PYTHONIOENCODING=utf-8`，否则 Windows 终端中文乱码。
- 禁在长命令尾部挂 `| tail`/`| head`（管道缓冲会丢日志）；长命令用日志重定向 + Read 工具读。
- 本机 VPN 在注册表写系统代理且 `ProxyOverride` 为空，会把本地 127.0.0.1 回环请求劫持给代理（返回 502）。本地回环必须直连（`client.py` 的 `ProxyHandler({})` 已处理）；出站请求（如 `git clone`）必要时显式 `export https_proxy=<系统代理地址>`（地址/端口查注册表 `Internet Settings\ProxyServer`）。

## 工程开发纪律

### 分支纪律（最高优先级）

- **所有功能/修复/文档开发一律先建分支，禁止直接在 main 上提交开发代码**。改库名、加功能、改依赖、写文档、改 CI，都是「开发」，都要走分支。
- 分支命名：`feat/<短横线名>`（新功能）、`fix/<短横线名>`（bug 修复）、`docs/<短横线名>`（文档）、`build/<短横线名>`（构建/CI/依赖）。
- 建分支前确认 main 干净：`git status` 无未提交改动 → `git checkout main` → `git checkout -b <分支名>`。
- 唯一允许 main 直改的例外：**紧急修复**（CI 挂掉、build 破损、阻塞性 bug），且提交时在对话里说明为何没走分支。
- 合并纪律：分支自测全绿（`ruff check` / `ruff format --check` / `mypy` / `pytest`）后才合并。合并方式 `git checkout main && git merge <分支> --no-ff` 或走 PR；合并后删除已合并分支（`git branch -d <分支>`）。
- 多任务并行时每个任务独立分支，互不干扰。

### Office 窗口清理（每次验证后，必做）

- **每次测试/验证通过后，主动关掉拉起的 Office 窗口**：`uv run offipy quit ppt`（/excel/word）。注意 `quit` 只是 `obj.Quit()`，PowerPoint 进程可能残留（加载项/关闭对话框卡住）——quit 后**确认进程已退**：`tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"`，仍有残留则 `taskkill //F //PID <pid>`，**清理后再确认一次，不留任何 Office 进程**。
- 改了 server 依赖的模块（如 `ppt.py`、`server.py`、`mcp_server.py`）后，必须重启 8890 的 server 进程（`taskkill` 该 PID 后 CLI 自动重建），否则 server 加载旧代码。
- **收工检查不看「是否用了 COM」**：就算本轮只跑纯转换（convert.py / Playwright 渲染），也可能有之前会话拉起的 Office 窗口没关。**每轮验证收尾，一律 `tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` 确认零残留，不猜**。2026-08-04 教训：冒烟只跑 convert.py，想当然以为没拉 PowerPoint，实际 POWERPNT.EXE 残留 415MB；`offipy quit` 后仍卡加载项，需 `taskkill //F //PID`。

### 资源型开发（资产库 / 图标库 / 素材）

- 建立艺术库、图标库、字体/素材类资产时，**先 `WebSearch` 调研 + `gh` 查 GitHub 现成方案**
  （开源库、许可证、维护状态、成熟度），别闭门造车。选型后记 ADR 与来源，尊重许可证
  （MIT / Apache / SIL OFL 等），资产文件注明出处。

### 提交纪律

- 小步提交：指定具体文件、清晰 message，前缀 `feat:` / `fix:` / `docs:` / `build:` / `chore:` + 一句话说明。
- 提交前先 `git status` / `git diff` 确认只含本任务改动，不带无关文件。
- 当前主干 `main`，但 main 只接受合并结果与紧急修复。

### 文档同步纪律（每次修完代码 / 发布新版本后，必做）

- **代码修复与版本发布都要及时更新文档，不能只改代码**。行为变化（哪怕语义相同）
  必须同步到 `docs/`、README、CHANGELOG，避免文档与代码脱节。
- **收尾检查清单**（做完一轮代码改动后逐项过）：
  - `Glob **/README*` 排查所有 README，核对版本号、特性列表、使用示例、结构图是否跟上；
  - `Grep` 旧版本号与受影响的关键词（如被改动的 API 名、行为描述、错误码），确认无残留；
  - 改了 schema（OpSpec）后必须重跑 `scripts/gen_api_stub.py` / `scripts/gen_api_ref.py`，
    api.pyi 与 docs/api/* 由脚本派生，手改无效；
  - 破坏性/行为变化 → 更新 `docs/migration.md` 迁移指南；
  - CHANGELOG 顶部版本必须匹配 `__version__`（`test_changelog_top_version_matches` 兜底）。
- **文档更新独立成 docs commit**（`docs: <一句话说明>`），不与代码/版本 bump 混在一个 commit。

### 开发文档不进 git 追踪

- **内部开发产物一律不入库**：本地开发文档统一放 `docs/development/`（研究/审计，如
  gap_analysis、ppt_research、ppt_design_research），计划放 `docs/superpowers/`——
  两者都在 `.gitignore`。
- 只有对外发布文档才进 git 追踪：`docs/index.md`、`docs/usage.md`、`docs/exceptions.md`、
  `docs/api/`（schema 生成）、`docs/protocol.md`、`docs/compatibility.md`、
  `docs/deprecation.md`、`docs/benchmarks.md`。
- 新增内部开发文档 → 放 `docs/development/` 或 `docs/superpowers/`（已在 ignore 内），
  不要 `git add`；误入库的用 `git rm --cached` 移出（文件保留在磁盘）。

### PyPI 发布纪律

- **PyPI 的发布只能由 GitHub Action Release 链触发**（quality → office-real → publish-testpypi → testpypi-smoke → gh-release → publish-pypi 全绿后自动发布），**禁止手动 twine 上传或任何绕过 CI 质量门禁的强制发布**。手动发布会让发布链无法干净重跑（版本已存在即拦）且丢失质量门禁（v0.11.0 手动发布后需手动补 gh-release 的教训）；发布一律走 tag → push → Release workflow。

## 版本号更新规则

- **单一来源**：版本号只写在 `src/offipy/__init__.py` 的 `__version__`；`pyproject.toml` 经 `[tool.hatch.version]` 自动读取，**别处不重复写**。
- **语义化版本（SemVer）** `MAJOR.MINOR.PATCH`：
  - `MAJOR`：不兼容的 API 变更（改包名、改调用签名、删既有操作）
  - `MINOR`：向后兼容的新功能（新增原子操作 / CLI 子命令 / 管线特性）
  - `PATCH`：bug 修复、内部重构、不改变对外行为的小调整
- **0.x 开发期约定**：破坏性变更升 `MINOR` 即可，不升 `MAJOR`（尚未对外承诺稳定性）。
- **每次升版本**：单独成一个 commit（message 形如 `chore: bump version to 0.2.0`）；若已发布 PyPI，同步打 tag `v0.2.0` 并推送。功能 commit 不带版本号，升版本也不夹带功能改动。
- **验证**：改完跑 `uv run python -c "import offipy; print(offipy.__version__)"` 确认生效。
- **每次升版本后必查配置与文档**：版本迭代（MINOR/MAJOR 必查，PATCH 视影响面）后，先用 `Glob **/README*` 排查所有 README，再逐一核对是否需要同步更新：
  - `pyproject.toml`：依赖增减、sdist/wheel 的 include/exclude 是否覆盖新增目录（如新增 `examples/`、`docs/`）
  - 根 `README.md`：特性列表、使用示例、结构图、MCP 配置示例是否跟上新功能
  - 各子目录 README（`src/offipy/_vendor/`、`examples/` 等 vendored/示例说明）
  - `examples/` 新增示例里出现的命令要真实可跑（冒烟过再发布）
  需要更新就随版本一起修（单独 docs commit），并确认无残留旧版本号/旧命令。
