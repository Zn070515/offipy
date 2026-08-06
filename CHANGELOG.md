# Changelog

格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本语义遵循 SemVer
（正式首发前版本首位恒为 0，破坏性变更只升 MINOR）。

## [0.12.0] - 2026-08-06

### Added
- **offipy.art 艺术分析子系统**：纯标准库、确定性、**只建议**的视觉/排版质量分析——`build_scene`
  把幻灯片抽象成 ArtScene，5 个维度规则（层级 / 构图 / 排版 / 颜色 / 媒体）评估；grade /
  confidence / evidence_coverage 三分离，证据不足维度降级 `insufficient_evidence` 不误报，无总分
  门禁（取舍留给调用方）
- **双源场景合并**：measurements（浏览器像素证据：颜色/字号/文本）+ pptx（几何审计快照）按
  「文本强佐证 + 几何兜底」一对一匹配；未匹配元素保留 + warning，绝不静默丢弃
- **组合入口**：`analyze_deck(pptx=..., measurements=..., profile=...)` 一次调用几何审计 + 艺术分析；
  `deck.render_with_quality_report(html, audit_mode=..., fail_on=..., profile=...)` HTML→PPTX 生成即
  质量参考（返回 `QualityRenderResult`）
- **内置 profile**：`balanced` / `consulting` / `academic` / `technology` / `event`
  （`profile_names()` 可查、`get_profile(name)` 可读可扩展）
- **基线对比 v2**：`compare_reports(before, after)` 产出 `ArtReportDiff`（finding 新增/消失/变化 +
  grade 变化），只建议不阻断

### Notes
- **零额外依赖**：`import offipy` 不加载 python-pptx / AI / COM；`import offipy` 即有 art 能力。
- **证据诚实边界**：只传 `pptx=` 时 hierarchy / typography / color 维度证据不足 → 自动降级
  （`insufficient_evidence` + `art.evidence.limited` warning），纯几何规则照常运行。
- **RenderedSlide（PNG / slides_dir 像素级）分析延后到 v0.12.1**：`build_scene(slides_dir=...)` 明确拒绝。
- 完整规则 / 证据源 / 边界见 [`docs/art.md`](docs/art.md)。

## [0.11.6] - 2026-08-06

### Added
- **Ppt `close_pres` 新操作**（#26）：API 补齐 close 语义——`ppt.close_pres(save, doc_id)`
  关闭演示文稿，配合 `deck.open_live` 释放 PowerPoint 对产物的锁，同路径 re-render
  不再 PermissionError
- **只读 op 支持 `follow_active`**（#25）：`get_cell` / `read_range` / `read_doc_text` /
  `read_slide_texts` / `read_slide_summary` 在会话语义下可跟随活动目标，
  MCP / CLI / Remote 同步放行

### Fixed
- **deck 混合 flex/非 flex 布局页**（#21）：naturalDisplay 按 slide 记录，flex 页不再被
  全局 block 强压成两栏竖排
- **deck `open_live` 源产物不再被锁**（#22）：打开前复制到 `offipy-live-*` 临时副本，
  `close_live` 释放；目标被 PowerPoint 占用时给可操作错误（替代裸 PermissionError）
- **Word 错误路径前置校验**（#23/#28/#30）：`set_table_col_width` / `add_list` /
  `format_text` / `insert_image` / `open_doc` 非法输入在进入 COM 前转
  `InvalidArgumentError`，不再裸抛 COM 错误
- **Excel 错误路径前置校验**（#24/#31）：`set_range` 维度不匹配 / `merge_cells` 非法区域 /
  `open_book` 源文件缺失同理
- **Ppt 错误路径前置校验**（#27/#32）：`add_slide` 非法 layout / slide 索引越界 /
  `add_picture` 源文件缺失同理
- **CLI 退出码一致**（#29）：`deck make` / `deck outline` 缺必填参数统一 exit 2
  （对齐 argparse 参数错误语义）

## [0.11.5] - 2026-08-06

### Fixed
- **Word 合并表格列宽真正生效**（使用方复测 #15）：0.11.4 的逐格回退仍走
  `Columns(col).Cells`，而混合列宽表格上 `Columns(col)` 本身即被拒访（「表格有
  混合的单元格宽度」），回退首句就抛错、从未生效。改为逐行 `table.Cell(r, col).Width`
  设宽（r 遍历 `Rows.Count`），彻底绕开列对象
- **`split-2col` 布局 `.cols` wrapper 形态回归修复**（使用方复测 #9）：
  - `display: contents` 的 `.cols` 无盒（`getBoundingClientRect` 恒 0×0）但子元素
    照常渲染，measure 的 `isHidden` 不再把它当隐藏剪掉子树——`.cols` wrapper 形态
    不再产出 ZERO shapes（比 0.11.4 的「两栏堆叠左列」更糟的回归）
  - 容器 `align-items` 由 `stretch` 改为 `flex-start`，只让 `.col` 用
    `align-self: stretch` 填满整页；眉标/标题等非 col 的 flex 子元素保持自然高度，
    不再被拉成整页高的 7.8in 超高文本框

## [0.11.4] - 2026-08-06

### Fixed
- **Office 进程残留 / 僵尸实例治本**（真实使用项目 13 项 Issue 反馈 #8–#20 反哺）：
  - `quit()` 按实例 PID 精确回收：调用 `Quit()` 后轮询进程退出，Excel/Word/PowerPoint
    常驻残留（RCW/COM server 保持）超时未退则 `taskkill /F` 精确清理本实例——不按
    进程名模糊清理，不误杀用户其它 Office 窗口。`quit()` 返回值恢复契约：正常退出
    返回 `None`，仅「实例已死」返回 `True`（不再误报失败）
  - server `_rebuild` 重建连接前先 `reap_own_process` 清掉本库附着过的僵尸进程，
    治「外部 kill 后重连附着僵尸实例，后续 op 反复 OLE error 0x800ac472」
  - core 新增 `app_process_pid`（COM 窗口句柄反查 PID）/ `wait_process_exit` /
    `reap_process`，`quit_app` 同步按 PID 清理
- **Word / PowerPoint 交互修复**：
  - 合并单元格后 `Columns(col).Width` 拒访（「表格有混合的单元格宽度」）→ 首次回退
    逐格设宽（0.11.5 修正：该回退在混合列宽表格上仍被拒访、从未生效，见 0.11.5 条目）
  - `read_doc_text` 归一化 Word 原始标记：表格 cell 结束符 `\x07` → `" | "`、
    段落符 `\r\n` / `\r` → `\n`（`\r\n` 先行避免双换行），Agent 文本层读到干净结构
  - `add_picture` 内嵌图片传 `SaveWithDocument=msoTrue(-1)`（LinkToFile=False 时传 0
    被 PowerPoint 拒为 E_INVALIDARG）+ 路径 normpath 归一化
  - `insert_image` 插图落文末 `Range`，不再覆盖既有内容
- **deck / 审计 / 审美 / 布局 / CLI**：
  - `deck render` 后审计目录从随机 `tmp` 名改名为最终输出名保留（`<stem>_audit`），
    aesthetic/feedback 回路按最终名自动发现测量数据，不再丢失
  - `charts` 图表容器匹配：转换器 `measure.py` 的 shape 记录补 `className`，按
    `chart` token 命中（`chart-note` 仍排除）
  - audit text-fit 缺码位（如 Arial 无 CJK 码位）记 0 宽 → 按字符权重兜底非零宽 +
    低置信标注，长文不再被误放行
  - 半透明颜色（`rgba`）按 source-over 与底色合成后再算对比度；全透明按「无颜色」
    回退页面背景——透明黑不再被当不透明黑造成假阳性
  - `split-2col` 布局 `.cols` 包装层 `display: contents` 透明化兜底 + 模板注释提醒
    flex-item 约束，Claude 多包一层 wrapper 也能正确两栏
  - CLI `--key` 全局归一化 `-` → `_`（`--doc-id` 与 `--doc_id` 等价，README 双写法
    不再打架）；`--payload` 数组报错补 list 用法提示

## [0.11.3] - 2026-08-06

### Fixed
- **audit text-fit / overlap 度量治本**（真实 deck `wastevision_deck.pptx` 第二轮反馈反哺）：
  - `text.fit.vertical` 读取段落 `a:lnSpc` 行距（`spcPts` 绝对 / `spcPct` 百分比），未设才
    回退 `字号×1.2`；行尾软换行（`a:br`）不再多算一行空行
  - `text.fit.horizontal` 用 fontTools 解析字体（`.ttf`/`.ttc`），按 hmtx 字宽 + kerning/GPOS
    字距求和，对齐 PowerPoint DirectWrite 度量（此前 Pillow 不带字距，拉丁粗体约高估 1.3%）
  - `geometry.overlap.partial` 新增 `decorative_layering` 豁免：双方无文本的长条装饰分层
    （短边≥3 倍、较长一维 ≥70% 大者、非包含）部分重叠不再误报
- **转换器盒高与宽度冗余治本**：
  - 显式分行（`<br>`）文本框盒高按 `行数×行高` 兜底，消除「行数×行高 − 亚像素取整」的
    系统性偏短（约 1%）
  - 单行标签统一加 2.5% 宽冗余（此前仅 CJK 单行，拉丁+粗体仍欠度）

## [0.11.2] - 2026-08-06

### Fixed
- **audit text-fit / covered_text 误报治本**（真实 deck 使用反馈反哺）：
  - text-fit 横溢**仅对显式 `wrap="none"` 的文本框**报——`square`（含 bodyPr@wrap 未设，
    PowerPoint 默认自动折行）永不报横溢；段落含 `a:br` 软换行时取**最长段**宽，不再跨段
    求和（旧版把软换行各段宽度加总，5 段 11pt 文本曾算出 16.25in 的假横溢）
  - 超宽 / 超高加 **1pt 噪声下限**——Pillow FreeType 与 PowerPoint DirectWrite 度量引擎
    亚 pt 级差异不再触发误报
  - 字体定位支持 **`.ttc` 集合**（微软雅黑 `msyh.ttc` 等），消除字体找不到时的大面积
    「字符估算低置信」回退
  - `covered_text` **只在被盖形状有文本时发射**——空框 / 装饰点浮空卡片（双方无文本）
    不再报「完全覆盖」
- **转换器单行 CJK 框加 2.5% 宽冗余**：HTML→PPTX 渲染时含 CJK/全角字符的单行
  `wrap="none"` 文本框加宽冗余，抵消浏览器与 PowerPoint 的字体度量差异，防溢出不可换行

## [0.11.1] - 2026-08-06

### Fixed
- **overlap/margin 误报治本**（真实 deck 使用反馈反哺）：extract 提取显式填充
  （`a:noFill` 透明），overlap 改为统一遮挡判定——透明无文本上层不遮挡
  （`transparent_overlay`）、实心小装饰浮有文本容器（`decorative_overlay`）、
  文字浮无文本背景（`text_on_background`）分别豁免；全宽贴边条归
  background/header/footer 豁免 margin（`full_bleed` / `header_footer`）。
  有文本上层（透明与否）一律不豁免——文本叠文本仍报

## [0.11.0] - 2026-08-06

### Added
- **PPTX 静态质量审计**（`offipy audit` / `audit_pptx`，见 docs/audit.md）：纯解析 `.pptx`
  （ZIP+XML），**不开 PowerPoint、不依赖 Microsoft Office**，静态几何门禁检查越界 / 贴边 /
  重叠 / 文本溢出 / autofit 风险；输出稳定 `rule_id`（机器键）与 text / json / markdown / html
  报告（html 为单文件内联 SVG 画布，可按严重度筛选，可选 PNG 页面背景）；`--fail-on` 按严重度
  门槛阻断，退出码 0/1/2/3 不与其它命令的 `OffipyError → 1` 冲突；`import offipy` 不加载
  python-pptx（惰性 import 硬约束）
- **基线回归**（`compare_pptx` / `offipy audit --baseline ... --fail-on-new`，见
  docs/audit-baseline.md）：形状匹配链（同 shape_id → name+type → 归一化文本 hash+几何邻近 →
  图片 sha256）聚合新增/已解决/变化的问题与形状增删移动缩放文本变化；`--fail-on-new` **只阻断
  候选新增或恶化**的问题，基线历史问题放行
- **Deck 生成门禁**（`deck.render_with_report` / `deck make --audit-mode strict --fail-on ...`）：
  HTML→PPTX 渲染后即审计，达 `fail_on` 抛 `AuditGateError`（报告先落盘、临时文件已清理、
  旧 `.pptx` 不被破坏）；`render()` 签名与行为不变
- 固定验收集：`tests/fixtures/audit/`（synthetic / edge_cases / baseline / candidate /
  deck_generated）驱动逐规则断言与误报控制（connector / hidden / rotate / flip / group 不误判）
- 文档：docs/audit.md + audit.en.md + audit-baseline.md + audit-baseline.en.md（挂 mkdocs nav）

## [0.10.2] - 2026-08-05

### Fixed
- `quit(force=True)` 传不进（真实使用反馈）：`_Facade.quit` 硬编码无参版本遮蔽了 `PptApp/WordApp/ExcelApp` 的 `quit(force)`——连到既有 Office 实例时错误消息引导用户传 `force=True` 却抛 `TypeError`（死路）。现在 `quit(force=False)` 显式透传，`Excel()/Word()/Ppt()` 三 facade 一并修复
- 误导性报错：未知 `doc_id` 的 `TargetNotFoundError` 不再写「用 list_docs 查看当前打开的」——`doc_id` 是**会话内**标识，本地直连 `Ppt()/Excel()/Word()` 与会话式 `Remote*/CLI/HTTP` 互不相通，`list_docs()` 显示在也不代表本会话能查到（真实使用：`Ppt().open_pres` 的句柄喂给 `deck.export_slides` 误导排查）。报错现在点明「当前会话」+ 跨会话边界
- work-copy 静默切换（真实使用）：`render()` 首次生成 `.audited.html` 副本后改源 HTML 不生效——convert 静默复用旧副本。现在源 HTML 比副本新（mtime 更大）时自动重建副本，改源即刻生效
- 图片缺失静默占位图（真实使用）：`<img src="fig/xxx.png">` 引用文件缺失时不报错，converter 静默嵌入 3-4KB 空白占位图（几何审计无差别，只能靠 ppt/media 文件大小暴露）。现在 measure 检测 `naturalWidth/Height=0` 破图，convert fail-fast 报错列出缺失文件

## [0.10.1] - 2026-08-05

### Fixed
- `read_slide_summary` title 回退漏过滤空文本 shape（真实比赛 PPT QA 发现）：纯文本框 deck 上 header 背景矩形带空 TextFrame（top=0）按阅读顺序排在真标题前时，title 拿到空串。现在 title 回退与 body 一致，跳过空文本候选（对齐 `read_slide_texts` 的 `include_empty=False` 语义）。

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
- 连接级 503（`_reply_503`）竞态：`shutdown(SHUT_RD)` 挡不住已进接收缓冲的请求字节，
  close 时缓冲非空触发 TCP RST，客户端偶发读到 `ConnectionResetError` 而非 503——
  改为非阻塞 `recv` 读光缓冲再 close（正常 FIN），修复 flaky 的 `test_concurrency_limit_503`
  （CI windows 3.13 偶发失败）

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
