# Changelog

格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本语义遵循 SemVer
（正式首发前版本首位恒为 0，破坏性变更只升 MINOR）。

## [Unreleased]

### Fixed
- 连接级 503（`_reply_503`）竞态：`shutdown(SHUT_RD)` 挡不住已进接收缓冲的请求字节，
  close 时缓冲非空触发 TCP RST，客户端偶发读到 `ConnectionResetError` 而非 503——
  改为非阻塞 `recv` 读光缓冲再 close（正常 FIN），修复 flaky 的 `test_concurrency_limit_503`
  （CI windows 3.13 偶发失败）

## [0.10.0] - 2026-08-05

### Added
- `read_slide_texts` 重构为按页读 shape 文本：`read_slide_texts(slide_idx, *, include_empty=False, recursive=True)` 返回 `list[SlideTextRecord]`（TypedDict：shape_id/name/text/坐标/占位符类型/group_path），递归 group 内文本；`read_slide_summary()` 承接旧的全部页摘要语义
- 公共数据模型冻结：`SlideTextRecord` + `PLACEHOLDER_TYPE_NAMES`（完整 PpPlaceholderType 映射，`unknown_{n}` 兜底），`from offipy import SlideTextRecord` 可用
- 发现性：`dir(Excel()/Word()/Ppt()/Remote*)` 显示 schema 全部正式 op + `quit`（与 CLI/MCP/文档同批）
- 固定 fixture：`tests/fixtures/ppt/minimal_text_shapes.pptx` + 生成脚本入库；P2-2 fixture 结构验证测试（python-pptx 重开，CI 无 python-pptx 时跳过）
- 迁移指南：`docs/migration.md`（0.9 → 0.10）
- P1-6 类型门禁：api.pyi stub 生成器保留 keyword-only `*`；read_slide_texts 签名快照测试 + mypy 用户示例 reveal_type 测试

### Changed
- **Breaking**：`read_slide_texts()` 签名破坏性变更——旧的无参全部页摘要用法改调 `read_slide_summary()`；`slide_idx` 保持必填，缺失抛标准 `TypeError`（方案 A，无运行时拦截，长期签名不被污染）
- 摘要豁免集（P1-2）：页码/页眉/页脚/日期占位符（type 13/14/15/16）+「页码候选」不进 title/body；纯文本框页面改为稳定阅读顺序的启发式摘要（排序稳定、语义一致，不承诺逐字节与 0.9 相同）
- 占位符常量修正（P0）：`PP_PLACEHOLDER_TITLE=1`、`PP_PLACEHOLDER_CENTER_TITLE=3`（0.9 错标 13/14，实为 slideNumber/header），新增 `PP_PLACEHOLDER_SLIDE_NUMBER/HEADER/FOOTER/DATE`——`set_title`/`set_body`/`set_notes` 在标准布局上行为更正确

### Fixed
- P0 占位符常量 Bug（见 Changed）
- `read_slide_texts` 对纯文本框 deck 返回空（旧实现只读 HasTitle/Placeholders(2)）——现按文本能力读全部 shape，含 group 内文本

## [0.9.0] - 2026-08-05（正式首发）

> **发布策略**：TestPyPI 预发布用 `0.9.0a1` / `0.9.0rc1` 编号，稳定发布时 `__version__`、git tag、
> CHANGELOG 顶层三者必须一致（对齐测试兜底）。**0.9.0** 为正式首发（继承 0.9.0a1 全部已验证内容），
> 外部用户验证稳定后再升 **1.0.0**（`__version__`/CHANGELOG 顶层同步）。

### Added
- 高层 API facade：`offipy.Excel() / Word() / Ppt()` 上下文管理器，未显式定义的 op 代理到底层 app（`api.py`）
- 会话语义：三件套 op 优先解析实时 `ActiveDocument / ActiveWorkbook / ActivePresentation`，仅失败时回退缓存句柄 + liveness probe
- offipy 异常体系：`OffipyError` 基类 + 领域异常（`InvalidArgumentError / TargetNotFoundError / FileConflictError / ComOperationError / ProtocolError / OfficeUnavailableError / ServerStartError / RemoteCallError / ConversionError / UnsupportedPlatformError`），均带 `code`，Python / RPC / MCP 三入口同源
- 跨平台惰性 COM import：`import offipy` 在非 Windows 上可用，调用 Office API 时抛 `UnsupportedPlatformError`
- server 安全：Bearer token 鉴权（env-first + 文件持久，token 写失败即启动失败）、`/status` 健康端点、16MB body 限制（413）、Content-Type 校验（415）、op 白名单、鉴权 `/shutdown` 优雅停机
- 目标身份原语：每 App `get_target()` 返回 active 文档身份（app/name/path）；`/status` 报告当前目标；只读不创建（无打开文档时读 op 抛 `TargetNotFoundError`，不再隐式 `Add()`）
- destructive 目标绑定：破坏性 op 支持 `expected_target`（doc_id / name / path 可组合，resolve-once）精确比对，绑定失败拒绝执行，防切焦点误改；`activate()` 同步真实 UI
- operation schema 单一来源：新增 `schema.py` 声明式 op 表，server `_OPS` / CLI 参数类型 / MCP 工具注册全部由 schema 派生——新增 RPC 只改一处
- 线程前端 + 单 COM worker 队列：慢 COM op 不阻塞 `/ping` / `/status` / `/shutdown`；COM 对象仅 worker 线程触碰（套间安全）
- `OperationResult` 统一返回契约：`{ok, operation, resource_id, message, data}`（附 `result` 兼容别名），原始 COM 对象不外泄
- 有界资源 + 幂等：client 超时对齐 server 600s；`request_id` 幂等缓存（LRU 512 / TTL 600s，超时重试不重执行）；COM 队列上限 64、并发线程上限 16，满则 503 busy
- `save` / `save_pdf` 覆盖保护：目标存在且 `overwrite=False` 时抛 `FileExistsError`
- save/close 防弹窗：`save()` / `close_book()` / `close_doc()` 对从未保存的文档自动落盘 `<cwd>/<名字>_<时间戳><ext>` 并返回绝对路径，不再弹「另存为」对话框
- converter vendored 进 wheel（`src/offipy/_vendor/`）+ chromium 预检（渲染前检查浏览器，缺失给安装提示）
- deck 原子替换：渲染写同目录临时文件 + `os.replace`（`mkstemp` 随机名，并发不互踩），失败不破坏既有 .pptx；`render` / `make` 输出已存在且 `overwrite=False` 时抛 `FileExistsError`
- CLI 复杂参数系统：重复 `--key` 聚合为 list、`--payload/--json` JSON 透传、未知参数报错退出码 2；`--port` 子命令继承父级
- `offipy server status|stop|restart` 子命令：`/status` 真实握手 + 进程管理（PID 文件含 port/pid/token_sha256/started_at 归属验证 + netstat 探测）
- Agent 只读 op：`word read_doc_text` / `ppt read_slide_texts` / `excel read_range` 读回文本层
- extras 拆分：`office / deck / mcp / all`，核心 `import offipy` 零额外依赖；`py.typed` 随 wheel 分发（PEP 561）
- License（PEP 639）：`license = "MIT AND ISC"` SPDX 表达式（vendored 转换器 / Phosphor 图标为 MIT，内嵌 Lucide 为 ISC），`license-files` 随产物分发
- 治理文档：CHANGELOG / SECURITY / CONTRIBUTING / THIRD_PARTY_NOTICES；README 重写（品牌 offipy、安全模型、会话语义）

### Changed
- 命令行更名：`office` → `offipy`（console script 只留 `offipy`）；品牌统一（对外文案 / 文档 / 示例页脚统一 `offipy`）
- `mcp` 依赖收窄为 `>=2.0,<3.0`；`pyproject.toml` 补 `project.urls` / `keywords` / `classifiers`；authors 邮箱换 GitHub noreply（公开元数据不暴露学校邮箱）
- server 改继承 `ThreadingHTTPServer`：每请求独立线程，COM 只由 worker 队列线程消费；非回环 host 默认拒绝（`--unsafe-allow-remote` 显式放行）
- DisplayAlerts 作用域化：构造器不再永久静音，op 内自存自还（修复嵌套 wrapper 互踩）；三件套 DisplayAlerts 全局副作用退出时还原
- deck CLI 真布尔：`--html/--out/--no-open/--feedback/--theme/--layouts/--overwrite` 改专用 argparse 选项（`store_true`），根除 `bool("false")` 翻真 bug
- Word / PPT 导出 PDF 改走官方 `ExportAsFixedFormat`
- 客户端 HTTP 错误语义化：网络 / 超时 / 非 JSON 一律抛 `RemoteCallError`
- CLI 参数按目标签名类型转换（`_coerce_kwargs`），弃全局值猜测
- MCP 工具结构化返回（void op → `ok (op)`，有值 op 原样透传）；`save` / `save_pdf` 透出 `overwrite`
- 库层不再抛 `SystemExit`（CLI 层捕获 offipy 异常并以退出码 1 报告）
- 公开文档统一对齐当前技术栈：README / docs/ 如实描述三入口返回形状（`OperationResult` 为 HTTP-only 契约）；`docs/api/` 由 `schema.py` 重新生成；新增英文版（`README.en.md` / `SECURITY.en.md` / `CONTRIBUTING.en.md` / `docs/*.en.md` / `docs/api/*.en.md`）

### Fixed
- PPT DisplayAlerts 常量改正（`ppAlertsNone=1`，原 `=0` 实为 `ppAlertsAll`）
- HTTP 边界：`do_POST` 非 `/call` / `/shutdown` 路径 404（此前任何路径都当 /call）；负 `Content-Length` 拒绝 400；响应体超上限回 500 不写大 payload
- deck CLI `--overwrite false` 此前被 `bool("false")` 翻成 True 误覆盖目标文件
- `_parse_cell` 畸形坐标（如 `"A1B2"` 被字母/数字拆分误读）与越界坐标（列 > XFD、行 > 1048576）此前会被误解析，现已收严并抛 `InvalidArgumentError`
- Linux 纯模块兼容：user_data_dir 路径分隔符跨平台断言、envcheck 依赖平台感知（非 Windows 不查 pywin32 等）、版本断言改 PEP 440

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
- MCP server：三件套操作暴露为 MCP 工具
