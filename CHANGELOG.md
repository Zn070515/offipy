# Changelog

格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本语义遵循 SemVer
（正式首发前版本首位恒为 0，破坏性变更只升 MINOR）。

## [Unreleased]

## [0.9.0] - 2026-08-04（未发布）

> **预发布编号策略**：正式首发 1.0.0 前，TestPyPI 预发布用 `0.9.0a1` / `0.9.0rc1`
> 编号；稳定发布时 `__version__`、git tag、CHANGELOG 顶层三者必须一致（对齐测试兜底）。
> 0.9.0 从未发布，不重复 bump——round-2 修复后版本仍为 0.9.0。

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

### Fixed
- PPT DisplayAlerts 常量改正（`ppAlertsNone=1`，原 `=0` 实为 `ppAlertsAll`）

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
