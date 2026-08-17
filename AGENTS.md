# AGENTS.md

> 给 AI 编程代理（Codex / Claude Code 等）的 offipy 仓库交接说明。
> 本文件是工程协作契约，代理接手前必须完整读一遍；与 `CLAUDE.md` 内容一致，二者任一都足够独立工作。

## 项目是什么

**offipy**：Windows-only 的 Office COM 自动化库，会话式驱动 Word/Excel/PowerPoint，外加 HTML→可编辑 PPTX 管线。

- 核心目标：让 AI 独立产出**美观、符合审美、言之有物**的 Office 产物（PPT 为主）。
- CLI 命令 `offipy`，`import offipy` 跨平台可跑（调 Office API 才需 Windows）。
- 仓库：`https://github.com/Zn070515/offipy`，包名即 `offipy`。

## 环境要点（务必遵守）

- **Windows + Git Bash**（命令一律 bash 语法）。PowerShell 不是默认壳。
- **用 `uv` 管理 `.venv`**。系统 `python` 是 3.10，项目按 `.python-version`（**3.12**）走 venv。
  所有 Python/CLI 命令必须用 `uv run` 前缀：
  - `uv run python ...` / `uv run pytest ...` / `uv run offipy ...`
  - `uv run mkdocs build --strict`（docs 构建）
- **调用 Python 子进程前 `export PYTHONIOENCODING=utf-8`**，否则 Windows 终端中文乱码（GBK）。
- **禁在长命令尾部挂 `| tail` / `| head`**：管道缓冲会丢日志，长命令用日志重定向
  `命令 > logs/xxx.log 2>&1`，之后用 Read 工具读日志看结果。
- **代理**：本机 VPN 在注册表写系统代理且把本地回环劫持给代理（返回 502）。
  - 本地 127.0.0.1 回环**必须直连**（`client.py` 的 `ProxyHandler({})` 已处理）；
  - 出站请求（`git clone`、`gh`、下载）必要时显式
    `export https_proxy=http://127.0.0.1:12334`（端口查注册表 `Internet Settings\ProxyServer`，可能变）。

## 质量门禁（分支合并前必须全绿）

```bash
export PYTHONIOENCODING=utf-8
uv run ruff check .
uv run ruff format --check .
uv run mypy src/offipy
uv run pytest tests -q
```

- ruff：`line-length=100`，规则集很宽（含 A/ARG/C4/PERF/PT/PYI/TRY 等），`TRY003` 特意关闭。
- mypy：`strict = true`，`platform = win32`（Linux CI 也按 win32 检查，别改成 Linux 平台）。
- pytest：`-ra --strict-markers`、`filterwarnings = error`（任何 warning 都算失败）。
  注册 marker 只有 `com`（需活 Office/COM server）和 `deck_render`（需 chromium）——**写测试前核对 marker，别用未注册的**。
- 全量门禁约 3000+ passed / 40+ skipped。本地 Windows 全绿 ≠ CI Linux 纯模块绿——发布前必须看 CI。

## 分支与提交纪律（最高优先级）

- **所有功能/修复/文档开发一律先建分支，禁止直接在 main 上提交开发代码**。
  分支命名：`feat/<名>`（新功能）、`fix/<名>`（bug）、`docs/<名>`（文档）、`build/<名>`（CI/依赖）。
- 建分支前确认 main 干净：`git status` 无未提交改动 → `git checkout main` → `git checkout -b <分支名>`。
- main 只接受**合并结果**与**紧急修复**（CI 挂、build 破损、阻塞性 bug），紧急修复要在对话里说明为何没走分支。
- 合并：分支自测全绿后 `git checkout main && git merge <分支> --no-ff`，合并后 `git branch -d <分支>`。
  **不要 amend**（GitHub Actions 依赖 commit 历史，amend 会破坏 tag 对齐）。
- **小步提交**：指定具体文件、清晰 message，前缀 `feat:` / `fix:` / `docs:` / `build:` / `chore:` + 一句话。
  提交前 `git status` / `git diff` 确认只含本任务改动。

## 版本管理

- **单一来源**：版本号只写在 `src/offipy/__init__.py` 的 `__version__`；`pyproject.toml` 经 hatch 自动读取，别处不重复写。
- **SemVer** `MAJOR.MINOR.PATCH`；0.x 开发期破坏性变更只升 MINOR（未承诺稳定性）。
- **每次升版本独立一个 commit**（`chore: bump version to 0.x.y`），功能 commit 不带版本号，bump 不夹带功能。
- 升版本后**必查**：`Glob **/README*` 排查所有 README 版本锚点、`CHANGELOG.md` 顶层版本号。
  - `tests/test_packaging.py::test_changelog_top_version_matches` 断言 CHANGELOG 首个版本标题 == `__version__`；
  - `test_readme_version_anchors_match` 断言 README ×2 的「当前版本」锚点——**bump 后不同步会红**。
- 验证：`uv run python -c "import offipy; print(offipy.__version__)"`。

## 发布纪律（PyPI 只能走 CI 链）

- **禁止手动 `twine upload` 或任何绕过 CI 质量门禁的发布**。手动发布会让发布链无法干净重跑且丢门禁。
- 发布流程：**bump 版本 → 确认 main 最近 CI 全绿（`gh run list`）→ 打 tag `vX.Y.Z` → push tag** →
  `.github/workflows/release.yml` 自动跑
  `quality → office-real → publish-testpypi → testpypi-smoke → publish-pypi → gh-release` 全绿后自动发布 PyPI。
- release.yml 的 office-real 走**自托管 runner**（标签 `[self-hosted, windows, office]`），真机 COM 必须真实安装 wheel 后跑。

## Office 进程清理（每次验证后必做）

- 每次测试/验证通过后主动关 Office 窗口：`uv run offipy quit ppt`（/excel/word）。
- `quit` 只是 `obj.Quit()`，PowerPoint 进程可能残留（加载项/关闭对话框卡住）。
- **收尾一律 `tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` 确认零残留，不猜**；有残留则 `taskkill //F //PID <pid>`，清理后再确认一次。
- 就算本轮只跑纯转换（convert.py / Playwright 渲染）也可能有之前会话拉起的 Office 窗口——每轮收尾都查。

## Server 会话纪律（改服务端代码必读）

- 架构是「常驻 8890 HTTP server + COM worker」：`server.py` 载代码，`cli.py`/`mcp_server.py` 只发 RPC。
- **改了 server 依赖的模块后必须重启 8890**：`offipy server restart`（或 taskkill 该 PID 后 CLI 自动重建），
  否则 server 加载旧代码。**新功能参数报 `unexpected keyword` / 保护逻辑不生效 = server 陈旧的高危信号**。
- `offipy server status` 看 `version` 是否 == 当前 `__version__`；不一致就是僵尸 server。
- 本机直连 facade（`offipy.direct.Excel/Word/Ppt`）与会话式 `Remote*/CLI/HTTP` 是**两个物理隔离的 doc 注册表**，
  跨会话引 doc_id 报 TargetNotFoundError 是设计如此，报错已说明「doc_id 只在同会话内有效」。

## 文档同步纪律（每次修完代码/发版本后必做）

- **不能只改代码**。行为变化必须同步到 `docs/`、README、CHANGELOG，避免文档与代码脱节。
- **文档更新独立成 `docs:` commit**，不与代码/版本 bump 混在一个 commit。
- 改了 schema（OpSpec）必须重跑 `scripts/gen_api_stub.py` / `scripts/gen_api_ref.py`——api.pyi 与 docs/api/* 由脚本派生，手改无效。
- 破坏性/行为变化 → 更新 `docs/migration.md`。
- 内部开发文档（`docs/development/`、`docs/superpowers/`、`development-history/`）**不入库**（gitignored 且 mkdocs `exclude_docs` 排除）。

## 架构速览（改代码前先定位）

```
src/offipy/
  core.py         COM 应用生命周期与会话管理（惰性 COM、GetActiveObject、liveness 探针）
  server.py       常驻 HTTP server（会话式跨调用保活，Bearer token，幂等 request_id，16MB/64MB 上限）
  client.py       会话式客户端（Remote*/CLI/MCP 共用，ProxyHandler({}) 直连回环）
  excel.py/word.py/ppt.py   三套件原子操作（会话驱动，doc_id 权威）
  direct.py       本地直连 COM facade（Excel()/Word()/Ppt()，STA 套间）
  cli.py          offipy 命令入口（exit 2=参数/预运行错，exit 1=运行时领域错）
  mcp_server.py   MCP server
  schema.py       op 声明式单一来源（OpSpec）——新增 RPC 只改这里
  deck.py         HTML→可编辑 PPTX 管线（render/make/audit，Playwright 实测 DOM 坐标）
  design.py/layouts.py/autopick.py/aesthetic.py   设计系统（token/布局/自动选型/审美审计）
  feedback/       可学习反馈系统（numpy MLP，学习路径只读指定目录）
  charts.py/icons.py/animations/   原生图表/图标/动画后处理注入
  _vendor/        vendored 第三方（html_to_editable_pptx 转换器、diagram-design）
```

关键架构原则：
- **核心 `import offipy` 零额外依赖**：pywin32/python-pptx/lxml 等一律**函数内惰性 import**
  （顶层 import 会破坏「纯 COM 用户 import offipy 即炸」契约，有测试锁死）。
- 图表/图标/动画都是「convert 照常渲染占位 → 后处理读 measurements.json（className + rect 坐标）→
  python-pptx 替换成原生对象」的注入架构；vendored 转换器不可改。
- `server.py` 顶层不 import pythoncom/pywintypes；COM 单 worker 线程（套间安全）。

## 高频踩坑区（改这几类代码前必看）

**COM / Office 对象模型**
- HRESULT 负值显示用两补码 `f"0x{hr & 0xFFFFFFFF:08x}"`；`0x800401F0`（CO_E_NOTINITIALIZED）= 当前线程从未 CoInitialize，不是「COM 未运行」。
- pywin32 early-binding 用 `gencache.EnsureDispatch`，`%TEMP%\gen_py` 缓存损坏会报 `no attribute 'CLSIDToClassMap'`——删缓存目录自愈。
- PowerPoint 占位符类型：`ppPlaceholderTitle=1`（不是 13！13 是 slideNumber）、`ppPlaceholderCenterTitle=3`。按类型找占位符，别硬编码 `Placeholders(2)`。
- Word `write_line` 末尾留空尾段 → 刚写入文本恒在 `Paragraphs.Count - 1`，`add_heading` 样式要落在 `Count - 1`，否则标题样式上到空尾段、目录为空。
- Excel 从未保存的脏工作簿 `Close` 会弹「另存为」卡死 COM 线程——Close 前先 `book.Saved = True`。
- early-bound COM 布尔读回 -1：断言 `in (True, -1)`，别 `is True`。
- 中文版 Excel 新建工作簿默认名是「工作簿N」不是「BookN」——别按英文名硬编码，用句柄/`list_docs()`。
- python-pptx 没有的属性不会报错只会**静默 no-op**（`BaseShape` 用 `__setattr__` 兜底），样式改动必须实机渲染核对。

**Windows / 环境**
- 中文乱码：Python 子进程前 `export PYTHONIOENCODING=utf-8`；`subprocess.run(text=True)` 必须 `encoding="utf-8", errors="replace"`。
- 本地回环被 VPN 劫持（502）：直连不走代理。
- 别用 `Path(x).is_file()` 判「路径 or 内容」——超长字符串 Linux 抛 ENAMETOOLONG、Windows 返回 False，要 try 住 OSError。
- 测试里别写 Windows 硬编码路径（`C:\...`），用 `tmp_path`——否则 Linux CI 的 posixpath 语义下预检不触发、红得莫名其妙。

**CI / GitHub Actions**
- 自托管 runner 必须**从 Git Bash 会话拉起**（`cmd //c 'C:\actions-runner\run.cmd'` 全路径），否则 `shell: bash` 解析到 WSL 的 System32 bash 全挂。
- reusable workflow 的**调用方 job 不能写 `timeout-minutes`**（报「workflow file issue」0 jobs 秒挂无日志），只能写在被调用 workflow 内部 job 上。
- `cache-dependency-glob` 是字符串标量不是 YAML 数组。
- actionlint（ci.yml 的 workflow-lint job）校验全部 `.github/workflows/*.yml`，改 workflow 后本地跑
  `uv run actionlint` 或看 CI workflow-lint job，防「YAML 结构错 → 0 jobs 秒挂」。
- GitHub `env` key 大小写不敏感去重，别同时写 `http_proxy` + `HTTP_PROXY`。

## 文档站与基准

- mkdocs 文档站（docs/，双语 index/usage/api/...）→ GitHub Pages `https://zn070515.github.io/offipy/`，
  由 `.github/workflows/docs-deploy.yml` 在 push main 时自动发布。本地构建：`uv run mkdocs build --strict`（必须 --strict 绿）。
- 兼容矩阵只标 **Windows 11 + Microsoft 365 + Python 3.12 为 Tested**，其余 Expected（未实测），别乱改。
- 性能基准 `docs/benchmarks.md`、质量审计 `docs/audit*.md` 是快照文档，改了性能/质量行为后同步。

## 收尾检查清单（做完一轮改动后逐项过）

- [ ] `Glob **/README*` 排查所有 README，版本号/特性/示例/结构图是否跟上
- [ ] `Grep` 旧版本号与受影响关键词（被改的 API 名、行为描述、错误码）无残留
- [ ] 改了 schema → 重跑 `gen_api_stub.py` / `gen_api_ref.py`
- [ ] 破坏性/行为变化 → 更新 `docs/migration.md`
- [ ] CHANGELOG 顶部版本匹配 `__version__`
- [ ] 四门禁全绿（ruff check / format --check / mypy / pytest）
- [ ] `tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` 确认 Office 零残留
