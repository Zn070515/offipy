# Changelog

格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本语义遵循 SemVer
（正式首发前版本首位恒为 0，破坏性变更只升 MINOR）。

## [0.18.0] - 2026-08-15

### Added

- **feedback 学习质量 8 项（#115-#122）**：
  - tiny_image 特征补全 + FEATURES schema v1→v2 bump（#115）：媒体特征新增
    `art.media.tiny_image.area_ratio`；旧记录 feature_schema_version 不符 →
    不可训练样本（冷启动回退 v2 语义不变）
  - 预处理标准化（#118/#120）：训练集拟合**零方差特征 drop + 高相关（|r|≥0.99）
    去重 + 全局 z-score**，mean/scale/kept 持久化到 model.json `preprocessing` 块
    （model schema v2），推理端走同一 transform
  - 容量自适应 soft warning（#121）：容量按独立样本数自适应（Lunt & Xu H≈√n），
    `samples_per_param` 分级 ok/warn/critical，只记录不拒绝写盘
  - 样本级 repeated stratified CV（#119）：按 rule 分层的重复 5 折、**绝不做
    pair-level split**（Kajimura 2022）；95% 置信下界触 chance →
    `poor_generalization` soft flag，不拒绝写盘
  - insufficient_pairs 逐规则诊断（#117）：训练样本不足时返回 per-rule
    `{fixed, accepted, pairs, single_direction, suggest}` 可行动建议
  - ensemble K=5（#122）：多 seed member 取平均降方差 + worth 校准
    （quality_score 归一）+ abstain（|worth| 近零 / member 分歧大的 finding 不
    shift）+ OOD（特征 z 越界不 shift，保守回退 v2）
  - CLI deck audit 透传 experimental score（#116）：`deck audit --feedback-dir`
    时报告 / `--json` 输出带 `experimental_score`
  - `feedback status` 表面 `effective_dims` / `samples_per_param` /
    `poor_generalization`（#121/#119）

### Fixed

- **deck srcset 还原越界（#110）**：源 HTML 含字面 `__OFFIPY_DATA_N__` 占位符
  （N 越界）时不再抛裸 IndexError——未 stash 的占位符原样保留，交给 URL 重写按
  普通相对路径处理。

### Changed

- **mkdocs build --strict 纳入纯模块 CI 门禁（#108）**：任何文档构建告警即 CI 红
  （dev extra 新增 mkdocs + mkdocs-material）。
- **stdlib-random + 固定 seed 变异 fuzz（#110）**：drawio / HTML 解析器以固定 seed
  变异输入做鲁棒性 fuzz（含 URL / data 路径白名单验证）。

## [0.17.1] - 2026-08-15

### Fixed

- **feedback 学习 4 个修复（#111-#114）**：
  - 训练加数值门禁（#112）：loss 非有限 → `training_diverged`；输出恒定 → `model_collapsed`；全局梯度裁剪；坏模型不写，原子保留旧模型
  - 学习消费必须显式 `feedback_dir`（#113）：`feedback=True` 无 dir → `InvalidArgumentError`；`learned_adjustments` 无 dir → 回退 v2，不静默加载全局 `~/.offipy`
  - `severity_shift` 按规则证据门禁（#111）：仅 ≥3 有效标签的规则被 shift；`quality.score` 同语义（只由通过门禁的 finding 贡献）
  - CLI 打开学习消费通道（#114）：`deck audit --feedback-dir` 应用反馈学习；`deck make --export-png`（`--feedback` 为弃用别名）；`feedback append` 追加标签

## [0.17.0] - 2026-08-15

### Added

- 可学习 feedback 系统：自定义 numpy MLP（配对 margin loss + centering 先验），
  注册式输入输出（FEATURES 特征注册表 + OUTPUTS 输出注册表，均带 schema 版本）
- `offipy[feedback]` extra（numpy 必选）；核心 `import offipy` 仍零依赖
- CLI/MCP：`feedback_train` / `feedback_status`（schema app `feedback`）
- `art analyze` 学习路径：`rule.delta`（历史聚合 → feedback_severity_adjustments）、
  `finding.severity_shift`（后处理 pass，severity_override=False 才作用）、
  `quality.score`（替换 experimental_score，opt-in）
- 冷启动回退 v2（无模型/过期/无 numpy 时行为不变）；旧 v1/v2 API 完全向后兼容

## [0.16.2] - 2026-08-14

### Fixed

- **drawio 注入 3 个修复（#98-#100）**：
  - 正交/曲线边按折点（waypoints）渲染成折线，不再坍缩成单直线，保留箭头（#98）。
  - 多页 `.drawio` 在 deck 注入里**必须**用 `data-drawio-page="N"`（1 基）指定页，
    否则报错而不是静默取第一页（#99）；页名也可用（不区分大小写），`all` 不支持。
  - 节点/边的 `strokeWidth`、`rotation`、`dashPattern` 透传到 PPTX 形状
    （`dashPattern` 为空格分隔的数对，如 `3 2`）（#100）。

### Changed

- **CI 门禁加固（#101-#109）**：新增安全扫描 workflow（CodeQL / 依赖审计 / secret 扫描）、
  ruff 规则集扩展（补齐 A/C4/D/… 全量修复）、mypy 全量 `--strict`、coverage 门槛提升
  （Linux 纯模块 70 → 88）+ Windows 真机 coverage artifact、pytest `--strict-markers` +
  `filterwarnings=["error"]`、所有 CI action SHA-pin、job timeout + concurrency 防叠跑、
  RPC op 基准轻档纳入真机 CI（bench 脚本走显式 doc_id 路由 + slide_idx）。

## [0.16.1] - 2026-08-14

### Fixed

- **drawio 集成 5 个修复（#93-#97）**：
  - `data-drawio` 相对路径在 deck 管线内可用（#93）：HTML 被拷到临时目录时自动
    改写为 `file://` 绝对 URI，源文件仍能解析，不再报「drawio 源文件缺失」。
  - 空 `<div class="drawio">` 占位可测量（#94）：vendored measure.py 装饰门放行
    无背景无边框的占位，注入定位不再失败；容器测量尺寸为 0 时给可操作报错。
  - drawio 布局在容器内居中 + 文本随框缩放（#95）：非绑定轴留白对称，不再贴左上角。
  - 占位删除改几何一致匹配（#96）：只删与注入矩形同位置同尺寸的形状，不再误删
    bbox 内用户内容。
  - drawio `fontSize` 全链路提取与缩放（#97）：字号层级随容器缩放，不被拍平为 14pt。

## [0.16.0] - 2026-08-14

### Added

- **diagram app：diagram-design skill 的 LLM 设计能力以 Agent 原生模式接入**——
  offipy 自身不调 LLM、不 spawn agent：`diagram build` 把宿主 agent 按产物契约落地的
  Mermaid/drawio 源码文件转成可编辑 PPTX（16:9 整页）；`diagram install_skill` 把
  vendored diagram-design skill + offipy-diagram wrapper skill 安装到宿主 agent 技能
  目录（默认 `~/.claude/skills/`，幂等不覆盖，`--force` 覆盖重建）。
- 新增 CLI `offipy diagram build` / `offipy diagram install_skill`、MCP 工具
  `diagram_build` / `diagram_install_skill`，三入口经 schema 一处登记自动接线。
- `offipy-diagram` wrapper skill 钉死产物契约：Mermaid 仅支持 flowchart/graph、
  sequenceDiagram、stateDiagram-v2、erDiagram 四种 kind；其余类型改用 draw.io 表达。
- `diagram build` 输出覆盖保护对齐 Office 惯例：`out` 已存在时默认拒绝覆盖
  （`FileConflictError`），重新生成需 `--overwrite true`。

## [0.15.1] - 2026-08-14

### Fixed
- **server 脱敏与 503 排空（#81/#82/#83/#89）**：Windows 分支盘符路径边界补强
  （colon / 中文尾 / 英文与管道符边界），POSIX 分支按段特征收敛，修复路径
  泄漏与过度吞尾；`_reply_503` 排空缓冲加 1 MiB 上限，防恶意客户端持续灌字节
  卡死 accept 线程。
- **diagrams 分层与源判定（#84/#85）**：`_kahn_layers` 对环残留节点统一尾层、
  结果与遍历顺序无关（幂等）；`_read_source` 路径/文本二义判定——路径形态但
  文件不存在的字符串抛 FileNotFoundError 并给修正提示。
- **deck 出参校验（#86/#90）**：`export_slides` 与转换管线先验 out 参数、零页
  快失败，避免产物静默缺失。
- **doc_id 语义文档修正（#87）**：破坏性/需目标操作（close/save/save_pdf/
  add_sheet/export_slides 等）不再声称「缺省为活动」，改为必须显式传入 doc_id
  或 follow_active=True；schema、docstring、gen_api_ref 与 docs/api/* 全部同步。
- **gen_py 缓存损坏提示（#88）**：EnsureDispatch 撞 CLSIDToClassMap
  AttributeError 时抛 OfficeUnavailableError，并给出删除 %TEMP%\gen_py 的修复指引。

## [0.15.0] - 2026-08-14

### Added
- `offipy.diagrams.mermaid_to_pptx`：Mermaid flowchart → 可编辑 PPTX（16:9 整页，
  TD/TB/LR/RL/BT 方向，subgraph 容器，中文 label）。vendored diagram-design 的
  mermaid_extract.py 提取拓扑，offipy 自研分层布局 + python-pptx 渲染。
- deck 集成：HTML `<pre class="mermaid">` 块在渲染后被替换为可编辑形状
  （复用 charts 注入管线，需 visual audit 的 measurements.json）。
- 原生图示（draw.io）：`offipy.drawio.drawio_to_pptx(source, out_path, *, page=None)`
  把 `.drawio` 转成 16:9 可编辑 PPTX，保留作者版式与配色；deck 注入支持 HTML
  `<div class="drawio" data-drawio="file.drawio">` 声明。

### Fixed
- **server 脱敏盘符前紧贴 `\w` 漏脱（#80）**：Windows 分支 `\b[A-Za-z]:` 词边界在盘符前
  紧贴中文/英文/数字/下划线（无空格分隔）时失效 → 整条绝对路径原样泄漏。去掉 `\b` +
  盘符后 `[\\/]` 处 `(?!\/)` 排除 `http(s)://` scheme：无分隔拼接形态全脱，URL 不受影响。

## [0.14.5] - 2026-08-08

### Fixed
- **server 脱敏 file:// URL（#78）**：`file:///Users`、`file:///C:/` 等任意平台
  file:// 路径形态漏脱（0c02694 修 #75 时把 POSIX lookbehind 收紧，/ 前是 / 的形态
  全被挡死），现专用分支覆盖；http(s) URL 不受影响。
- **server 脱敏过度（#79）**：Windows 分支贪婪吞掉路径后尾随业务文本（「已保存
  C:\...pptx 到桌面」→ 尾随「到桌面」丢失）；POSIX 分支把 `../` 相对段误当绝对路径
  留残根。Windows 分支改非贪婪 + 业务文本边界截断（中文/引号/doc_id/串尾），
  POSIX 拒绝 `.` 前导相对段。已知限制：英文短尾词（done/processed）仍被吞（LOW，
  H10 宁多勿漏）。

## [0.14.4] - 2026-08-08

### Fixed
- **server 脱敏覆盖补全（#75）**：`_redact_message` 覆盖带空格的 Windows 绝对路径
  （`C:\Users\John Doe\...`）与任意 POSIX 根（`/data`、`/workspace` 等）——不再枚举根
  白名单、不再遇空格截断，`error`/`trace` 对任意形态绝对路径统一替换 `[REDACTED]`。
- **转换器 records 容器校验（#76）**：装配边界对 records 容器本身做可迭代校验，`records`
  为非可迭代标量（`42`/`true`）时降级为空列表，不再抛 `TypeError` 整体崩溃。
- **注入副本 URL 承载属性补全（#77）**：`<body background>` / `<blockquote cite>` /
  `<form action>` 等 URL 承载属性相对路径一并重写为绝对 `file://`；`data-*` 逻辑值不误伤。

## [0.14.3] - 2026-08-08

### Fixed
- **资产 accent rgb() 值域（#65）**：`rgb()`/`rgba()` 分量超界（>255）按 CSS 语义钳制，
  不再抛 `ValueError` 让主题 accent 兜底崩溃。
- **图标 SVG 解析守卫（#66）**：`_svg_to_subpaths` 接入统一 `parse_svg` 守卫，拒绝
  DOCTYPE/ENTITY（billion-laughs）并包装畸形 XML 为 `InvalidArgumentError`。
- **server 失败消息脱敏（#67）**：`error` / `trace` 消息内的绝对路径（Windows/POSIX/UNC）
  与 `doc_id` 值统一替换为 `[REDACTED]`，杜绝经响应泄露服务器信息。
- **转换器 record 结构校验（#68）**：装配边界过滤非 dict / 无 `kind` 的畸形 record，
  下游 `rec["kind"]` / `rec.get` 不再 `KeyError` / `AttributeError`。
- **注入副本 `data` 属性重写（#69）**：`<object data>` / `<embed data>` 相对路径一并
  重写为绝对 `file://`；`data-*` 自定义属性（data-icon/data-chart-data）不误伤。
- **audit 顶层尺寸容错（#70）**：损坏 `sldSz@cx/@cy` 非数字时幻灯片尺寸降级 0.0 +
  告警 `audit.extract.slidesize_corrupt`，不再整文件 audit 崩溃。

## [0.14.2] - 2026-08-08

### Fixed
- **deck 管线数据安全（分支A）**：审计目录归属保护 + `make`/`open_live` 泄漏清理、原子提交、
  TOCTOU 竞态、渲染超时杀进程、srcset 处理加固。
- **资产系统加固（分支B）**：primitives 文本控制字符拒绝、`_descending_pt` 分数步进修复、
  SVG DOCTYPE/entity 守卫 + viewBox 逗号、legacy data-icon URI 百分号编码、license 白名单
  运行时强制、render H8/H9/viewBox 上限、accent rgb() 解析。
- **COM 层安全（分支C）**：`quit` 前抑制弹窗防挂死（ppt/word/excel）、`open_book`/`open_doc`
  路径 abspath 规范化、`export_slides` 宽高上界校验、`freeze_panes` 整数校验、`add_sheet` 幂等。
- **server 协议加固（分支D）**：高熵 doc_id、slowloris socket timeout、inflight 超时释放、
  error_code→HTTP 状态映射、traceback 脱敏、远程明文传输警告。
- **audit 误报根治（分支E）**：text-fit 按 `a:br` 段感知取最长段宽、仅显式 `wrap="none"` 报横溢、
  支持 .ttc CJK 字体度量、covered_text 仅在下层有文本时报、每页元素上限 + 压缩炸弹守卫。
- **vendor 转换器边界硬化（分支F）**：不可信 measurement 输入 graceful 收敛（DOM-API 覆盖攻击
  注入 NaN/Inf/非数值不崩）、XML 非法控制字符清洗、embed_fonts/self_check/measure 数值字段
  str/float 化。

## [0.14.1] - 2026-08-08

### Fixed
- **图元 CJK 字形宽度度量（#56）**：`fit_font_size` / `fit_font_size_wrapped` 按字形加权
  （CJK/全角 1.0em、拉丁 0.55em），中文标签不再溢出测量矩形。
- **注入副本相对资源路径断裂（#57）**：临时注入 HTML 写入前把相对 `src` / `href` /
  CSS `url()` 重写为以源 HTML 目录为基准的绝对 `file://` URL。
- **SVG picture 无栅格回退（#58）**：主 `a:blip` 挂 PNG raster fallback（共享惰性
  Playwright 渲染），`asvg:svgBlip` 仍指矢量；无 Playwright 时降级纯 SVG。
- **图元 fill 参数静默失效（#59）**：`fill="accent"` 跟随图元最终 accent，label-pill /
  timeline-node 的 `fill` 公共参数驱动填充、未传回退 accent。
- **未定义主题变量误报格式错误（#60）**：空串 token 视为「未定义」，报错不再误导。
- **deck 宽高比元数据矛盾（#61）**：vendor 转换器显式声明 `<p:sldSz type="screen16x9">`。

## [0.14.0] - 2026-08-08

### Added
- **资源系统（Asset System v1）**：`asset://` 统一资源管线（provider → 确定性测量 →
  占位符注入 → 渲染落位），产出 PowerPoint **可编辑原生对象**而非位图。
  `AssetRef` / `AssetRequest` / `AssetRegistry` / `get_default_registry()` 抽象 +
  `asset://<provider>/<kind>/<name>?k=v` URI 语法（参数规范化、`%23` 十六进制色往返）。
- **procedural provider（8 个确定性纹理）**：wave / blob / dot-grid / square-grid /
  rings / topography / circuit / gradient-orb，纯参数确定性生成，支持
  `data-asset-placement="background"` / `decorative` 沉底与装饰落位。
- **primitives provider（8 个可编辑原生图元）**：quote-mark / section-number /
  label-pill / metric-badge / timeline-node / process-arrow / device-frame /
  browser-mockup，全部由原生形状 + 文本框构成，文字可编辑；`background` 落位显式拒绝。
- **`assets.json` 溯源清单**：visual-audit 渲染产出用法清单（provider / 许可证 /
  来源 / 落位），满足合规审计。
- **HTML 声明语法**：`data-asset` + `data-asset-param-*` + `data-asset-placement`，
  原生图元简写 `data-primitive` 语法糖。

### Changed
- **ph/lu 图标内部迁移到 provider**：`data-icon="<集>:<名字>"` 内部走
  `ph` / `lu` provider 与统一 asset 管线，对外行为不变。
- **通用确定性测量/占位符管线**：资源声明有稳定 ID（`asset-s<slide>-<seq>`），
  每次渲染结构一致。
- **`no_visual_audit` 前置检查扩大**：除 `data-chart` / `data-icon` 外，也拒绝
  `data-asset` / `data-primitive` 声明（fail-fast，资源注入依赖 measurements.json）。

### Compatibility
- 旧 `data-icon` 图标写法**无需改动**，无迁移步骤。
- `import offipy.assets` 仍为纯标准库表面，不加载 python-pptx / Pillow / playwright。

## [0.13.2] - 2026-08-07

### Fixed
- **#40 feedback.load_records 脏行跳过**：`dimension`/`action` 校验前移到读路径，
  JSON 合法但维度非法的坏行直接跳过，`dimension_weights()` 不再崩裸 KeyError。
- **#41 Excel 三个 op 补 `_parse_range` 预校验**：`set_number_format` /
  `add_conditional_format` / `autofit` 畸形 `range_addr` 先抛
  `InvalidArgumentError`，不再落到 COM 抛裸 `ComOperationError`，与 5 个兄弟 op 对称。
- **#42 公共 API 边界不泄漏裸内建异常**：`offipy.op()` 未知应用、`art.get_profile()`
  未知 profile 一律抛 `InvalidArgumentError`（OffipyError 子类），不再抛裸
  ValueError / KeyError。
- **#43 `add_heading` 越界 level 显式拒绝**：level ≠ 1/2/3 抛
  `InvalidArgumentError`，不再静默降级为 Heading 1。
- **#44 `add_page_number` 参数序修正**：OpSpec 改为
  `(alignment, color, size, doc_id, mode)`，`mode` 为 keyword-only；
  `docs/api/word`（含 .en）同步。
- **#45 server queue.Full 回滚唤醒合并等待**：owner 入队回滚时用
  `_complete_entry` 同步 result + event，非 owner 线程不再 merge 到永不完成的
  entry 挂 600s 返回误导性 504。
- **#46 `quit` 对称拒绝 `follow_active`**：与 `expected_target` 一致拒绝，不再
  静默消费（protocol.md「quit 不接受两者」）。
- **#47 `client.request()` docstring 对齐**：400 带可识别 error_code → 领域异常，
  不再声称一律归 RemoteCallError。
- **#48 oplog 轮转在跨进程锁内**：`_rotate` 移到文件锁作用域内执行，且跨进程锁
  落在旁路 `.lock` 文件上——Windows 下数据文件自身 fd 会挡 rename（WinError 32），
  并发进程持有 fd 时不再静默不轮转、日志可超 5MB。

### Notes
- 纯 PATCH 修复（#40–#48），无破坏性 API / CLI / 契约变化。

## [0.13.1] - 2026-08-07

### Fixed
- **#39 deck audit 预运行无效输入退出码**：缺失文件 / 非法 `--profile` 等无效输入统一映射
  exit 2（参数或输入错），不再误报 exit 1；`docs/usage` deck audit 契约表同步为
  `2=参数或输入错`。
- **#38 art low_contrast 像素归因假阳性**：组合 PNG 不再把遮挡 / 邻近图片区域主色归因为文本
  背景；元素级对比度一律用声明 / 有效背景（`effective_background`），像素仅保留
  `declared_not_found` 低置信提示，修复白色文本被图片遮挡时误报 HIGH「对比度 1.14」。

### Notes
- 纯 PATCH 修复（#38 / #39），无破坏性 API / CLI / 契约变化。

## [0.13.0] - 2026-08-07

### Added
- **PPT 形状读取与编辑（S1）**：新增 `read_shapes()`（严格定位 + 冻结的 `ShapeInfo` / 形状类型
  契约）与形状编辑操作——几何 / 文本 / 字体、填充 / 轮廓 / 可见性、删除 / z-order，跨
  Python API / RPC / MCP / CLI 暴露。
- **艺术分析规则级反馈学习 v2（S2）**：art 报告 schema 升 0.3，反馈经
  `feedback_severity_adjustments` 有界调整规则严重度并带来源溯源（provenance），
  `build_scene` / `analyze_deck` 可 opt-in 反馈。
- **主题感知图表注入 + deck audit CLI（S3）**：图表颜色按幻灯片主题变体派生（确定性调色板 +
  覆盖）；chart-dominant 布局预检更精确；新增 `offipy deck audit` 临时渲染审计工作流
  （未知 profile 友好报错）。
- **Word 跟进（S4）**：`add_page_number` 新增 append / standalone（left/center/right）三模式；
  数值行距在 API 各面安全接受；去除列表可见的尾部空项。

### Changed
- **CLI 错误契约对齐（S5）**：冻结 CLI 退出码——`InvalidArgumentError` → exit 2（使用 / 参数 /
  预运行无效输入）、`OffipyError` 系 → exit 1（运行时领域失败）；`audit` 保留 0/1/2/3、
  `deck audit` 保留 0/1 专属契约。库层 fail-fast（`ppt.open_pres` 缺失源文件 /
  `ppt.export_slides` 输出目录是文件 / Excel 畸形区域）；stderr 消息清洗不泄露 traceback。

### Notes
- v0.13 五个 S（S1-S5）全量合入，无破坏性 API 变更；迁移说明见
  [`docs/migration.md`](docs/migration.md)。

## [0.12.2] - 2026-08-07

### Fixed
- **#33 MCP `close_*` 返回注解**：`_RETURN_ANNOTATION` 补 `"str|null"`，`close_book` /
  `close_doc` / `close_pres` 从 `object` 修正为 `str`。
- **#34 版本偏斜自愈**：client `_probe()` 识别「协议匹配但版本不一致」的旧 server →
  `mismatch`，`ensure_server` 按 pid 归属自动重启；`server_status()` 偏斜时返回含
  `version` 的可读 dict（非 offipy 进程协议失配才返回 None）。
- **#35 Excel 畸形区域 fail-fast**：`set_range` / `read_range` / `set_border` 经
  `_parse_range` 集中校验，非法地址统一抛 `InvalidArgumentError`（不再穿透原始 COM 错误）。
- **#36 HRESULT 可读显示 + 线程契约提示**：负 HRESULT 以两补码 `0x…` 显示；`CO_E_NOTINITIALIZED`
  （0x800401F0）单列识别，报错提示 `com_apartment()`（P1-4 线程契约）。
- **#37 保存锁重试 + quit 不静默**：`save` / `save_pdf`（ppt/word/excel）经
  `save_with_lock_retry` 锁感知短重试，目标文件被占用时给可读 `ComOperationError`；
  `ppt.quit(force=True)` PID 解析失败时以 liveness 探针兜底，进程仍活则明确抛错。

### Notes
- 纯 PATCH 修复（#33-#37），无破坏性 API / CLI / 契约变化；迁移见
  [`docs/migration.md`](docs/migration.md)。

## [0.12.1] - 2026-08-06

### Added
- **slides_dir 像素级分析（第三证据源）**：`build_scene(slides_dir=...)` / `analyze_deck(slides_dir=...)`
  读逐页 PNG（`slide_<n>.png` + `_deck_info.json` 指纹）增强 ArtScene——页面级背景估计 / 调色板 /
  background_like_ratio，元素级声明颜色验证（`declared_verified` / `declared_not_found` /
  `center_fill_verified` / `complex_background`）。惰性 Pillow（`offipy[deck]`）。
- **实验性规则 `art.composition.background_like_area`**：页面级留白提示（联合条件：背景置信 ≥0.7 ∧
  均匀度 ≥0.7 ∧ 元素占用 ≤0.5 ∧ 非全幅图片），`conf ≤ 0.3` 不驱动降级。
- **维度 reliability 加权聚合**：`DimensionAssessment.reliability` = applicable 确定性规则
  reliability 按 coverage 加权均值（experimental 排除），`minimum_reliability` 仅供调试。
- **证据契约**：`ArtFinding` 带 `evidence_sources` / `evidence_reliability` / `evidence_method`，
  Markdown/HTML 展示，compare 时证据来源/方法变化判 `changed`；`low_contrast` 像素验证背景路径。
- **deck 像素集成**：`render_with_quality_report(pixel_analysis="off"|"best_effort"|"required",
  preserve_pixel_slides=False, slides_output_dir=None)`，PNG 先落 staging、全链成功后提交。

### Changed
- **ART_SCHEMA_VERSION / ART_REPORT_SCHEMA_VERSION → 0.2**：0.1 报告可被 0.2 读取（证据字段默认
  None）；compare 跨 schema 仅建议性对比 + warning。

### Notes
- **组合 PNG 的证据边界**：像素证据不归因到元素——不做溢出/真实遮挡判断；只做页面级背景证据 +
  元素级声明色验证。
- 完整规则 / 证据源 / 边界见 [`docs/art.md`](docs/art.md)。

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
