# Changelog

格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本语义遵循 SemVer
（正式首发前版本首位恒为 0，破坏性变更只升 MINOR）。

## [Unreleased]

## [0.9.0a1] - 2026-08-05（预发布）

> **预发布编号策略**：正式首发 1.0.0 前，TestPyPI 预发布用 `0.9.0a1` / `0.9.0rc1`
> 编号；稳定发布时 `__version__`、git tag、CHANGELOG 顶层三者必须一致（对齐测试兜底）。
> 当前处于 TestPyPI 预发布：round-3/round-4（ChatGPT_v3/v4 审核）全量修复后以 `0.9.0a1` 冒烟；
> 预发布验证通过后转正式首发（`__version__`/CHANGELOG 顶层同步）。

### Added
- 高层 API facade：`offipy.Excel() / Word() / Ppt()` 上下文管理器，未显式定义的 op 代理到底层 app（`api.py`）
- 会话语义（P1.2）：三件套 op 优先解析实时 `ActiveDocument / ActiveWorkbook / ActivePresentation`，仅失败时回退缓存句柄 + liveness probe
- offipy 异常体系：`OffipyError` 基类 + `OfficeUnavailableError / ServerStartError / RemoteCallError / ConversionError / UnsupportedPlatformError`
- 跨平台惰性 COM import：`import offipy` 在非 Windows 上可用，调用 Office API 时抛 `UnsupportedPlatformError`
- server 安全：Bearer token 鉴权（env-first + 文件持久）、`/status` 健康端点、16MB body 限制（413）、Content-Type 校验（415）、op 白名单
- `save` / `save_pdf` 覆盖保护：目标存在且 `overwrite=False` 时抛 `FileExistsError`
- converter vendored 进 wheel（`src/offipy/_vendor/`）+ chromium 预检（渲染前检查浏览器，缺失给安装提示）
- CLI 复杂参数系统：重复 `--key` 聚合为 list、`--payload/--json` JSON 透传、未知参数报错退出码 2
- 治理文档：CHANGELOG / SECURITY / CONTRIBUTING / THIRD_PARTY_NOTICES；README 重写（品牌 offipy、安全模型、会话语义、审核修复清单）
- server 安全加固（round-2）：非回环 host 默认拒绝（`--unsafe-allow-remote` 显式放行）、token 写失败即启动失败（杜绝假活）、显式 RPC 白名单（会话内部方法不暴露）、惰性 COM import
- `offipy server status|stop|restart` 子命令：`/status` 真实握手 + 进程管理（PID 文件 / netstat 探测）
- Agent 只读 op：`word read_doc_text` / `ppt read_slide_texts` / `excel read_range` 读回文本层
- deck 覆盖保护：`render` / `make` 输出已存在且 `overwrite=False` 时抛 `FileExistsError`（fail-fast）
- 依赖拆 `convert` extra（HTML→PPTX 转换器依赖）+ `tomli`（3.10 支持），核心依赖瘦身

#### round-3（ChatGPT_v3 审核，2026-08-04）
- 领域异常体系（策略 A）：`InvalidArgumentError / TargetNotFoundError / FileConflictError / ComOperationError / ProtocolError` 全部继承 `OffipyError` 且带 `code`；RPC `error_code` 与异常一一对应，Python/RPC/MCP 三入口同源
- 目标身份原语：每 App `get_target()` 返回 active 文档身份（app/name/path）；`/status` 报告各 App 当前目标（只读缓存快照）
- 只读不创建（P0-8）：无打开文档时 `active_*` 纯探测返回 None，读 op 抛 `TargetNotFoundError`——不再隐式 `Add()` 造空文档
- destructive 目标绑定（P0-7）：破坏性 op 支持 `expected_target`（name/path）精确比对，绑定失败拒绝执行，防切焦点误改
- operation schema 单一来源（P1-2）：新增 `schema.py` 声明式 op 表，server `_OPS` / CLI 参数类型 / MCP 工具注册全部由 schema 派生——新增 RPC 只改一处
- 线程前端 + 单 COM worker 队列（P1-1）：慢 COM op 不阻塞 `/ping` / `/status` / `/shutdown`；COM 对象仅 worker 线程触碰（套间安全）
- `OperationResult` 统一返回契约（P1-3）：`{ok, operation, resource_id, message, data}`（附 `result` 兼容别名），原始 COM 对象不外泄
- 鉴权 `/shutdown` 优雅停机端点 + `server stop` 分状态处理：token 失配拒绝、端口被占且无法证明归属拒绝强杀
- `server status` 只读（P0-3）：未运行不隐式拉起；公开 `server_ready()` 只读探测
- extras 拆分（P1-5）：`office / deck / mcp / all`，核心 `import offipy` 零额外依赖；`py.typed` 随 wheel 分发（PEP 561）
- deck 原子替换（P0-5）：渲染写同目录临时文件 + `os.replace`，失败不破坏既有 .pptx
- CI 矩阵（P1-6）：pure-module（Linux + 覆盖率门槛）/ windows（3.10–3.13）/ wheel-smoke / office-real（自托管真机 COM + deck_render）
- MCP in-memory 协议层测试（P1-7）：无子进程直连 MCPServer 验证工具注册 / 返回 / 错误映射

#### round-4（ChatGPT_v4 审核，2026-08-05）
- 目标语义显式化（P0-4/5/6）：`expected_target` **resolve-once**——`doc_id`/`name`/`path` 三键可组合，空对象/未知键直接拒绝，解析出的 doc_id 注入方法参数（杜绝「校验 A 执行 B」）；`activate()` 同步真实 UI（`Workbook.Activate()` / `Document.Activate()` / PPT 窗口），失败回滚抛 `ComOperationError`；`resource_id` 改用 doc_id；`list_docs` 只报已登记句柄并含 `active`；`get_target(doc_id=)` 显式查询
- 有界资源 + 幂等（§4/5）：client 超时对齐 server 600s；`request_id` 幂等缓存（LRU 512 / TTL 600s，超时重试不重执行）；COM 队列上限 64、并发线程上限 16，满则 503 busy；PID 文件含 `port/pid/token_sha256/started_at` 归属验证；token 落盘权限 0o600
- deck 原子渲染加固（§7）：`render` 临时文件改 `mkstemp`（同输出目录随机名，并发渲染不互踩）；失败清理临时文件，绝不破坏既有 .pptx；源 HTML 缺失抛 `InvalidArgumentError`
- 发布门禁（P0-2/3）：`release.yml` 门禁链 quality → office-real（真机必过）→ gh-release → publish-testpypi → publish-pypi（OIDC Trusted Publishing）；`ci.yml` 的 office-real 成为 PR 合并门禁；`docs/release.md` 发布手册
- CLI / 转换边界收严（§6/11/12）：`offipy mcp` 缺 mcp extra 友好报错（exit 2）；必填参数调用前预校验（exit 2）；`_parse_cell` 收严 `^([A-Za-z]{1,3})(\d{1,7})$` 且越界 XFD/1048576 拒绝；`set_title`/`set_body` 空输入抛 `InvalidArgumentError`；CLI `--port` 子命令 SUPPRESS 继承父级；`/call` args 非 dict → 400；`offipy check --profile`；`install_smoke.py --profile`；`.github/dependabot.yml`
- License（PEP 639）：`license = "MIT AND ISC"` SPDX 表达式（vendored 转换器 / Phosphor 图标为 MIT，内嵌 Lucide 为 ISC），`license-files` 随产物分发
- CI 稳定（P0-1）：版本断言改 PEP 440（`packaging.Version`，兼容 a/b/rc 预发布）；`mypy` 固定 `platform=win32` 修复 Linux CI 误报 winreg 常量缺失；envcheck 依赖探测平台感知（非 Windows 不查 pywin32 等）
- save/close 防弹窗：`save()` / `close_book()` / `close_doc()` 对从未保存的文档自动落盘 `<cwd>/<名字>_<时间戳><ext>` 并返回绝对路径，不再弹「另存为」对话框；`overwrite` 覆盖保护 fail-fast（先于触 COM）

### Changed
- 命令行更名：`office` → `offipy`（console script 只留 `offipy`）
- `mcp` 依赖收窄为 `>=2.0,<3.0`
- `pyproject.toml` 补 `project.urls` / `keywords` / `classifiers`
- `.github/workflows/release.yml` 补全发布门禁（ruff → format → mypy → pytest → tag 版本校验 → build → twine → 安装冒烟 → `--verify-tag`）
- 库层不再抛 `SystemExit`（CLI 层捕获 offipy 异常并以退出码 1 报告）
- CLI 参数按目标签名类型转换（`_coerce_kwargs`），弃全局值猜测
- MCP 工具结构化返回（void op → `ok (op)`，有值 op 原样透传）；`save` / `save_pdf` 透出 `overwrite`
- Word / PPT 导出 PDF 改走官方 `ExportAsFixedFormat`；三件套 DisplayAlerts 全局副作用退出时还原
- 客户端 HTTP 错误语义化：网络 / 超时 / 非 JSON 一律抛 `RemoteCallError`
- CI 矩阵扩到 Python 3.10–3.13；deck 单测不真拉 chromium（浏览器前置检查 no-op）

#### round-3（ChatGPT_v3 审核，2026-08-04）
- server 改继承 `ThreadingHTTPServer`：每请求独立线程，COM 只由 worker 队列线程消费
- DisplayAlerts 作用域化（P0-4）：构造器不再永久静音，op 内自存自还（修复嵌套 wrapper 互踩）
- deck CLI 真布尔（P0-6）：`--html/--out/--no-open/--feedback/--theme/--layouts/--overwrite` 改专用 argparse 选项（`store_true`），根除 `bool("false")` 翻真 bug
- extras 由 `convert` 收敛为 `office / deck / mcp / all`；`convert` 能力并入 `deck`
- `offipy check` 依赖缺失提示按 extra 给出安装命令（office/deck/mcp）
- 品牌统一（P1-9）：对外文案 / 文档 / 示例页脚统一 `offipy`
- 非 Windows 收集修复（P1-6）：COM 测试模块改 fixture + skipif 短路，Linux 收集不炸

#### round-4（ChatGPT_v4 审核，2026-08-05）
- 公开文档统一对齐当前技术栈：README / docs/ 按方案 B 如实描述三入口返回形状（`OperationResult` 为 HTTP-only 契约）；`docs/api/` 由 `schema.py` 重新生成（close_book/close_doc/save 返回路径、`get_target` 带 doc_id、`list_docs` 含 active）；公开文档新增英文版（`README.en.md` / `SECURITY.en.md` / `CONTRIBUTING.en.md` / `docs/*.en.md` / `docs/api/*.en.md`，docs/api 英文由 `gen_api_ref.py` 翻译层生成，MCP/CLI 描述保持中文单源；mkdocs 导航补 English 区段 + 中英互链）

### Fixed
- PPT DisplayAlerts 常量改正（`ppAlertsNone=1`，原 `=0` 实为 `ppAlertsAll`）

#### round-3（ChatGPT_v3 审核，2026-08-04）
- HTTP 边界（P0-10）：`do_POST` 非 `/call` / `/shutdown` 路径 404（此前任何路径都当 /call）；负 `Content-Length` 拒绝 400；响应体超上限回 500 不写大 payload
- deck CLI `--overwrite false` 此前被 `bool("false")` 翻成 True 误覆盖目标文件

#### round-4（ChatGPT_v4 审核，2026-08-05）
- Linux 纯模块 CI 修复：user_data_dir 路径分隔符跨平台断言、envcheck 依赖平台感知、版本断言改 PEP 440
- `_parse_cell` 畸形坐标（如 `"A1B2"` 被字母/数字拆分误读）与越界坐标（列 > XFD、行 > 1048576）此前会被误解析，现已收严并抛 `InvalidArgumentError`

## [0.8.0] - 2026-08-04

### Added
- `offipy check` 环境就绪诊断：分组清单（Python / 依赖 / Office 三件套 / 浏览器 / server）+ `--json` 机器可读，只读不拉起应用

## [0.7.0] - 2026-08-04

### Added
- M5 Word 能力补齐：样式系统（`format_text` / `format_paragraph`）、页面结构（页眉页脚 / 页码 / 页面设置 / 目录）、列表与表格（`add_list` / 合并 / 边框 / 列宽 / 行高 / 自动调整）、文档辅助（`find_replace` / `insert_image` / `insert_page_break`）
- 上述能力全部暴露为 MCP 工具 + 可运行示例

### Fixed
- `add_heading` 标题样式错落到空尾段，致正文继承标题样式、目录为空（活 Word 实证修复）
- M5 表格边 top/left 映射校正 + `add_list` 首项漏符号

## [0.6.0] - 2026-08-04

### Added
- M4 Excel 能力补齐：合并单元格（`merge_cells` / `unmerge_cells`）、边框（`set_border`）、条件格式（cell 规则 / 数据条 / 色阶）、冻结窗格、打印设置、基础三件套（行高 / 数字格式 / 自动列宽）
- 上述能力全部暴露为 MCP 工具 + 可运行示例

### Fixed
- 条件格式 `between` 需 `value2`；databar 实机修正（`BarColor`）

## [0.5.1] - 2026-08-04

### Changed
- 图标渲染：stroke round 线帽 / join 对齐 Lucide 源、主题感知缺省色、`icons-row` 多行排布

## [0.5.0] - 2026-08-04

### Added
- M3 PPT 图标库：Phosphor + Lucide 双集 vendored、SVG path 解析与几何压平、freeform 矢量图标注入（非图片非贴图）
- outline `@icons` 指令 + 隐式 `icons-row` 布局

## [0.4.0] - 2026-08-04

### Added
- M2 PPT 原生图表：`charts` 模块（bar / line / pie 声明解析、measurements 定位、python-pptx 原生替换）
- outline `@chart` / `@chart-data` 指令；`deck.render` 转换后自动注入原生图表

## [0.3.0] - 2026-08-04

### Added
- M1 内容工作流：markdown 大纲 → 逐页结构化内容 → HTML 骨架（`deck outline` CLI、布局 autopick 推断、`@layout` 白名单校验）

## [0.2.1] - 2026-08-04

### Fixed
- COM 止血：三套件抑制模态对话框（DisplayAlerts）、Excel Close 用正确枚举、路径按调用方 CWD 绝对化、`quit` 不反拉起死实例、COM 对象序列化返回 null、错误响应回 500/404
- 设计系统 token 化；审美审计健壮性（一致性权重生效、畸形 measurement 防崩）
- CLI 细节：无子命令报 usage、bool/none 大小写不敏感

## [0.2.0] - 2026-08-04

### Added
- 设计系统全量：design token 模型 + 麦肯锡 / 学术 / 深色科技三主题、`aesthetic.audit()` 审美审计、`layouts` 命名布局库、`autopick` 自动选型、`feedback` 反馈学习
- `deck.render` / `deck make` 支持 theme 注入与 `apply_layouts`

## [0.1.0] - 2026-08-04

### Added
- 工程化基线：src-layout + CI + 工具链全绿
- vendored `html-to-editable-pptx`（vector-first HTML → 可编辑 PPTX 转换器）
- 会话式 COM server + CLI（Word / Excel / PowerPoint 原子操作）
- `ppt export_slides` 逐页导出 PNG 供视觉迭代；`deck.py` 编排 render → open_live → export_slides
- MCP server：三套件操作暴露为 MCP 工具
