# Contributing

offipy 是 Windows-only 的 Office COM 自动化库。欢迎提交修复与改进，但请先遵守下面的纪律——
它们是为了让每个提交都可审查、可回滚、不破坏正在进行的会话式自动化。

## 开发环境

- Windows + Git Bash。命令用 bash 语法。
- 用 `uv` 管理 `.venv`（Python 3.12）：
  `uv run python ...` / `uv run pytest ...` / `uv run offipy ...`
- 调用 Python 子进程前 `export PYTHONIOENCODING=utf-8`，否则 Windows 终端中文乱码。

## 分支纪律（最高优先级）

- 所有功能/修复/文档开发一律先建分支，禁止直接在 main 上提交开发代码。
- 命名：`feat/<短横线名>`（功能）、`fix/<短横线名>`（修复）、`docs/<短横线名>`（文档）、
  `build/<短横线名>`（构建/CI/依赖）。
- 建分支前确认 main 干净：`git status` 无未提交改动 → `git checkout main` → `git checkout -b <分支名>`。
- 唯一允许 main 直改的例外：**紧急修复**（CI 挂掉、build 破损、阻塞性 bug），提交时说明为何没走分支。
- 合并：分支自测全绿后 `git checkout main && git merge <分支> --no-ff`（或走 PR），合并后 `git branch -d <分支>`。

## 提交规范

- 小步提交：指定具体文件、清晰 message，前缀 `feat:` / `fix:` / `docs:` / `build:` / `chore:` + 一句话说明。
- 提交前 `git status` / `git diff` 确认只含本任务改动，不带无关文件。
- 版本号单一来源 `src/offipy/__init__.py` 的 `__version__`；升版本单独成 commit
  （`chore: bump version to X.Y.Z`），不夹带功能改动。
- 正式首发（1.0.0）前，一切版本首位为 0；破坏性变更升 MINOR，不升 MAJOR。
- **预发布编号策略**：正式首发前，TestPyPI 用 `0.9.0a1` / `0.9.0rc1` 等预发布编号；稳定发布时
  `__version__`、git tag、CHANGELOG 顶层三者必须一致（对齐测试兜底）。**未发布的版本号不重复
  bump**——0.9.0 正式版从未发布，预发布修复仍以 `0.9.0a1` / `0.9.0rc1` 迭代，不另升版本号。

## 门禁（合并前必须全绿）

```bash
uv sync --extra dev
uv run ruff check .
uv run ruff format --check .
uv run mypy src/offipy
uv run pytest tests -q
```

`src/offipy/_vendor/` 是 vendored 外协代码（HTML→PPTX 转换器），不做 lint/format/mypy 约束。

## 新增 RPC：只改 schema.py

`src/offipy/schema.py` 的 `OPS` 表是操作的**唯一注册源**——server 白名单（`server._OPS`）、
CLI 参数类型转换（`cli._coerce_kwargs`）、MCP 工具注册（`mcp_server._build_tool`）全部由它派生。
新增操作三步：

1. 在对应 App 类（`excel.py` / `word.py` / `ppt.py`）实现方法；
2. 在 `schema.py` 登记 `OpSpec`（`params` / `returns` / `readonly` / `destructive` / `description`）；
3. 三入口自动就绪——**不要**再手写 `server._OPS` 或 MCP decorator。

一致性由 `tests/test_schema_consistency.py` 兜底（schema ↔ `server._OPS` 双向一致、每个 schema op
都有对应 MCP 工具、readonly/destructive 标志一致）。

## 异常策略 A（领域异常）

库层失败一律抛 `OffipyError` 子类，**绝不再抛 `SystemExit` / 裸 COM 异常**：

- `InvalidArgumentError`（`invalid_argument`）：参数/输入非法（解析、常量表、范围校验）
- `TargetNotFoundError`（`target_not_found`）：没有打开的工作簿/文档/演示文稿，或 `expected_target` 绑定失败
- `FileConflictError`（`file_conflict`）：目标文件已存在且未显式 `overwrite`
- `ComOperationError`（`com_operation`）：App 方法内 COM 调用失败（保留 `hresult` 供断连识别）
- `ProtocolError`（`protocol`）：协议不符

每个异常带 `code`，与 RPC 失败响应 `error_code` 一一对应。CLI 在边界捕获转退出码 + stderr 提示；
MCP server 捕获后转工具错误返回；库调用方直接捕获即可。

## 验证 Office 集成后必做

每次测试/验证通过后，主动关掉拉起的 Office 窗口并确认进程已退：

```bash
uv run offipy quit ppt     # /excel /word；quit 只是 obj.Quit()，可能残留
tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"   # 仍有残留 → taskkill //F //PID <pid>
```

改了 server 依赖的模块（`ppt.py` / `server.py` / `mcp_server.py` 等）后，必须重启 8890 的
server 进程，否则 server 加载旧代码。

## server 测试注意事项

- `tests/test_server_security.py` 会起真实 HTTPServer（临时端口）验证鉴权 / 请求限制 / op
  白名单——不 dispatch 真实 op，因此**不需要 Office**。
- 401 不杀 server：token 校验失败只拒绝请求，server 继续服务 `/ping` 与正确 token 的调用。
- `offipy server status|stop|restart` 管理常驻进程（PID 文件 + netstat 探测）；改 server 代码后
  用 `offipy server restart` 重启，验证时确认 `tasklist` 无 Office 残留。
- `tests/test_mcp_server.py` 的 in-memory 测试不起子进程：用 mcp SDK 内存流
  （`create_client_server_memory_streams`）+ `_lowlevel_server.run` 直连协议层，验证工具注册 /
  结构化返回 / annotations / 错误映射——无 Office 也能跑，是 MCP 改动的首选回归。
- 非 Windows 收集：COM 测试模块的 `pytestmark` 用 `sys.platform != "win32" or not core.running(...)`
  短路 + `com` marker，Linux 上收集不炸、自动跳过。deck 转换测试挂 `deck_render` marker。

## 资源型开发（资产库 / 图标库 / 素材）

先 `WebSearch` 调研 + `gh` 查 GitHub 现成方案（开源库、许可证、维护状态、成熟度），
别闭门造车。选型后记 ADR 与来源，尊重许可证，资产文件注明出处。
