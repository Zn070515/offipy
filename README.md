# offipy

Live Microsoft Office automation via COM（会话式驱动）+ HTML-first 可编辑 PPTX 管线。
目标：让 Claude 能独立产出**美观、符合审美、言之有物**的 Office 产物（Word / PPT / Excel）。

- **库名 / 命令**：`pip install offipy`、`import offipy`、CLI 命令 `offipy`
- **当前版本**：0.9.0a1（TestPyPI 预发布；正式首发号 1.0.0）

## 特性

- **会话式常驻 server**：跨调用保持 Office 窗口存活、文档 / 工作簿 / 演示文稿状态不丢
- **三套件原子操作**：Word / Excel / PowerPoint 增删改 + 保存 / 导出 PDF
  - Excel 另含格式能力——合并单元格、边框、条件格式（cell 规则 / 数据条 / 色阶）、冻结窗格、打印设置、行高 / 数字格式 / 自动列宽
  - Word 另含版式能力——样式系统（文字 / 段落格式）、页面结构（页眉页脚 / 页码 / 页面设置 / 目录）、列表与表格（合并 / 边框 / 列宽 / 行高 / 自动调整）、文档辅助（查找替换 / 图片 / 分页）
- **实时文档会话语义**：op 默认作用在用户**当前激活**的文档上（ActiveDocument / ActiveWorkbook / ActivePresentation），关窗后自动重建
- **断连自愈**：用户关窗或 Office 退出后自动重建会话
- **HTML-first 管线 + 设计系统**：Claude 写 HTML 幻灯片 → 原生可编辑 `.pptx` → 实况展示 + 视觉迭代；内置设计 token、3 套主题、11 种布局、审美审计、自动选型、反馈学习（见下方「设计系统」）
- **MCP server**：把全部三套件操作暴露为 MCP 工具，Claude Desktop 等可直接驱动真实 Office
- **环境诊断**：`offipy check` 一键检查 Python / 依赖 / Office 三件套 / 浏览器 / server 是否就绪（`--json` 机器可读，失败退出码非 0）
- **server 进程管理**：`offipy server status|stop|restart` 用 `/status` 真实握手 + PID 文件 / netstat 探测管理常驻进程
- **Agent 只读回读**：`word read_doc_text` / `ppt read_slide_texts` / `excel read_range`
  把文档文本层读回（供 Agent 迭代），经 CLI / RPC / MCP 三路暴露
- **高层 API**：`offipy.Excel() / Word() / Ppt()` 上下文管理器，库内直接驱动（见下方「Python API」）

## 环境要求

- Windows + 已安装 Microsoft Office（Word / Excel / PowerPoint）
- Python ≥ 3.10（本仓库开发环境为 3.12）
- `import offipy` 在非 Windows 上可运行；调用 Office API 时抛 `UnsupportedPlatformError`
- 支持的平台 / Office / Python 组合矩阵见 [`docs/compatibility.md`](docs/compatibility.md)
  （Tested / Expected / Unsupported 三栏）

## 安装

```bash
uv venv --python 3.12 .venv
uv pip install -e ".[all]"            # 全部能力（office COM + deck 管线 + MCP）
uv run playwright install chromium    # 转换器依赖 chromium 做 DOM 测量
```

核心 `import offipy` 零额外依赖；按用途增量安装 extra：

- `.[office]`：Windows COM 自动化（Word/Excel/PowerPoint）
- `.[deck]`：HTML→可编辑 PPTX deck 管线（python-pptx / lxml / fonttools / playwright / Pillow）
- `.[mcp]`：MCP server（`offipy mcp`，Claude Desktop 等接入）
- `.[all]`：以上全部

转换器本体已 vendored 进 wheel，装完即可用；deck 管线需额外跑 `playwright install chromium`。

## 会话语义（读我）

`offipy` 是**会话式**驱动，不是一次性脚本：

- 首次调用自动在后台拉起常驻 server（`127.0.0.1:8890`）；之后所有操作打到同一进程。
- op 作用在**用户当前激活**的文档上——你在 Excel 里激活哪个工作簿，`set_cell` 就写到哪个。
  **不隐式新建**：没有打开文档时，需要文档的 op 抛 `TargetNotFoundError`，提示先
  `new_book` / `open_book`。用 `get_target` 查询当前激活目标身份：
  `offipy excel get_target` → `{"app": "excel", "doc_id": "book1", "name": "Book1", "path": "..."}`
  （无则 `null`）。`doc_id` 是会话内稳定标识，跨调用有效，不随改名变化。
- `activate(doc_id)` 把指定文档设为激活目标并**同步真实 UI**（Excel `Workbook.Activate()`、
  Word `Document.Activate()`、PPT 激活含该文档的窗口），同步失败回滚并抛 `ComOperationError`；
  `list_docs` 如实返回已登记句柄的 `{doc_id: {"name", "path", "active"}}`（不隐式枚举未登记的）。
- 破坏性 op 可带 `expected_target`（`{"doc_id": ...}` / `{"name": ...}` / `{"path": ...}` 可组合）
  做**目标绑定**：resolve-once——server 先按绑定键解析出目标 doc_id，再用解析结果执行操作；
  绑定失败抛 `TargetNotFoundError`——杜绝「校验 A 执行 B」，防止切到别的文档后被误改
  （Python client / RPC 层可用，绕开按焦点路由）。
- `quit excel` 等命令关掉应用；`__exit__`（Python API）**不**关 Office 窗口，窗口与文档跨调用保持存活。
- 用户手动关窗后，下次调用自动重建会话（断连自愈）。

## 安全模型

- server 只监听 `127.0.0.1`，不对外网开放。
- **Bearer token 鉴权**：启动时生成随机 token（环境变量 `OFFIPY_SERVER_TOKEN` 优先，否则持久化到
  `%LOCALAPPDATA%\offipy\token`）。所有请求需 `Authorization: Bearer <token>`，否则 401。
- op 白名单：只暴露各 app 类的公开方法；请求体上限 16MB；`Content-Type` 必须为 JSON；
  `POST` 只接受 `/call` 与鉴权 `/shutdown`，其余路径 404；响应体上限 64MB。
- **进程所有权（不杀策略）**：`server status` **只读**——未运行只报「未在运行」，**不隐式拉起**；
  `server stop` 优先走鉴权 `/shutdown` 优雅停机；token 失配（`auth_fail`）或端口被非 offipy 进程
  占用且无法证明归属（`mismatch`）时**一律不杀**，只提示手动处理。
- **别把 token 泄漏给不信任的进程**——拿到 token 即等于拿到你当前 Office 会话的读写权限。
- 详见 [`SECURITY.md`](SECURITY.md)。

## 使用

```bash
# 首次调用自动在后台拉起常驻 server；之后所有操作打到同一进程
offipy excel new_book
offipy excel set_cell --sheet 1 --cell A1 --value 100
offipy excel format_cell --sheet 1 --cell A1 --bold true --size 14 --bg "#38BDF8"
offipy excel merge_cells --sheet 1 --range_addr A1:B2
offipy excel set_border --sheet 1 --range_addr A1:D5 --side all --style continuous --weight thin --color "#D0D7DE"
offipy excel add_conditional_format --sheet 1 --range_addr C2:C5 --rule cell --operator greater --value 0 --bg "#C6EFCE"
offipy excel freeze_panes --sheet 1 --rows 1 --cols 0
offipy excel page_setup --sheet 1 --orientation landscape --fit_to_pages_wide 1
offipy excel set_number_format --sheet 1 --range_addr B2:B5 --fmt "#,##0"
offipy excel autofit --sheet 1 --range_addr A1:D5 --rows false

offipy word new_doc
offipy word write_line --text "你好，世界"
offipy word format_text --paragraph 1 --bold true --size 18 --color "#2251FF"
offipy word format_paragraph --paragraph 1 --alignment center --line_spacing double
offipy word set_header_text --text "季度报告"
offipy word add_page_number --alignment center
offipy word page_setup --orientation landscape --paper a4 --top_margin 60
offipy word insert_toc --levels 3
offipy word add_list --style bullet
offipy word merge_table_cells --table_idx 1 --start_row 1 --start_col 1 --end_row 1 --end_col 3
offipy word set_table_border --table_idx 1 --style single --color "#9AA5B1" --sides all
offipy word set_table_col_width --table_idx 1 --col 1 --width 140
offipy word find_replace --find 季度 --replace 半年度 --replace_all true
offipy word insert_image --path out/cover.png --width 360
offipy word insert_page_break

offipy ppt new_pres
offipy ppt add_slide --layout 2
offipy ppt set_title --slide_idx 1 --text "标题"

offipy check            # 环境就绪诊断：Python/依赖/Office/浏览器/server（--json 机器可读）
offipy server status    # 常驻 server 状态（/status 握手，只读不拉起）；stop / restart 同理
offipy excel get_target # 查询当前激活目标身份 {app,doc_id,name,path}（无则 null）
offipy word read_doc_text            # Agent 只读：全文档文本
offipy ppt read_slide_texts          # Agent 只读：逐页 title/body/notes
offipy excel read_range --sheet 1 --range_addr A1:B2   # Agent 只读：区域二维值
offipy quit excel
```

复杂参数用 `--payload '<json>'` 透传（覆盖同名 kwargs）；重复 `--key` 会聚合成 list。

## Python API

```python
from offipy import Excel, Word, Ppt

with Excel() as x:  # 上下文管理器；__exit__ 不关 Office 窗口（会话语义）
    x.new_book()
    x.set_cell(1, "A1", 100)
    x.save("out/report.xlsx")

with Ppt() as p:
    p.new_pres()
    p.add_slide(2)
    p.set_title(1, "标题")
```

未显式定义的 op 经 `__getattr__` 代理到底层 app；offipy 异常（`OffipyError` 家族）原样透传。

**返回契约**：三条入口（Python API / HTTP RPC / MCP）同源但**返回形状不同**，如实对照见
[`docs/api.md`](docs/api.md)：
- **Python API** 返回 App 方法原始值（`new_book`→`doc_id` 字符串、`get_cell`→单元值、
  `get_target`→dict …）；void op 返回 `None`；失败抛 `OffipyError` 领域异常。
- **HTTP RPC `/call`** 返回 `OperationResult`（**HTTP-only 契约**）：
  `{ok, operation, resource_id, message, data}`（附 `result` 兼容别名）。`operation` 是
  `"excel.set_cell"` 式全名；`resource_id` 如 `excel:book:book1` 标识本次操作作用的文档
  （`doc_id` 是会话内稳定标识，不用用户可改的 name）；`data` 是操作结果（读 op 原值，
  void op 为 `null`）。原始 COM 对象不外泄。
- **MCP 工具** 返回操作 `data` 载荷（读 op 原值；void op 为 `"ok (<op>)"`）。

**异常契约（error_code）**：失败统一为 `OffipyError` 子类，每个带 `code`：
`InvalidArgumentError`（`invalid_argument`）/ `TargetNotFoundError`（`target_not_found`）/
`FileConflictError`（`file_conflict`）/ `ComOperationError`（`com_operation`，保留 `hresult`）/
`ProtocolError`（`protocol`）。RPC 失败响应的 `error_code` 与异常一一对应，client 按表映射回
对应领域异常，三入口同源。

## 设计系统（HTML deck）

Claude 写 deck HTML 时，只需引用 design token（CSS 变量）并给每页打 `data-layout`，渲染时注入主题与布局，一份 HTML 换主题即换皮。示例见 [`examples/decks/design-system/deck.html`](examples/decks/design-system/deck.html)。

```html
<head>
  <style data-theme="mckinsey"></style>   <!-- 主题占位：render(theme=) 替换 -->
  <style data-layouts></style>            <!-- 布局占位：render(apply_layouts=) 替换 -->
</head>
<section class="slide hero" data-pptx-slide data-layout="hero-title">…</section>
```

- **内置主题**：`mckinsey`（咨询蓝）/ `academic`（学术极简）/ `dark-tech`（深色科技）—— 一套 token 覆盖，`design.theme_css()` / `design.inject_theme()`
- **命名布局库**：`hero-title` / `split-2col` / `cards-3` / `big-number` / `quote-frame` / `timeline` / `comparison` / `chart-dominant` / `icons-row` / `portrait-feature` / `closer` —— `layouts.inject_layouts()` 按引用注入
- **审美审计**：转换产出的 measurement → 留白比例 / 字号层级数 / 每页色数 / 对比度 / 跨页一致性打分，`aesthetic.audit()` 输出带分报告供迭代
- **自动选型**：`autopick.pick()` 从内容结构推荐主题 + 每页布局 + 理由（纯规则，可覆盖）
- **内容工作流**：把「言之有物」的骨架标准化——写一份 markdown 大纲
  （`# 标题` + 每个 `## 段` = 一页，`- ` 要点、正文、`@layout:` 指令），
  `offipy deck outline --input outline.md --out deck.html` 一键得到 HTML 骨架，
  布局自动推断，再 `offipy deck make --theme <主题> --layouts` 注入主题与布局成稿。
  示例见 [`examples/outline/quarterly-review.md`](examples/outline/quarterly-review.md)。

  ```bash
  offipy deck outline --input examples/outline/quarterly-review.md --out out/quarterly.html
  offipy deck make --html out/quarterly.html --theme mckinsey --layouts --out out/quarterly.pptx
  ```
- **原生图表**：图表区打 `data-chart="<类型>"`（`bar` / `line` / `pie`），数据放容器
  `data-chart-data` JSON 属性或页内 `<script type="application/json" data-chart-target="<选择器>">` 块，
  `deck make --layouts` 后自动替换成 PowerPoint 原生可编辑图表（双击即可改数据），不再是贴图。
  **前提**：图表所在页须引用 chart-dominant 布局（`data-layout="chart-dominant"`），容器才有可测的
  surface 矩形。示例见 [`examples/decks/charts/chart-demo.html`](examples/decks/charts/chart-demo.html)。

  ```html
  <div class="chart" data-chart="bar"
       data-chart-data='{"categories":["Q1","Q2"],"series":[{"name":"营收","values":[40,70]}]}'></div>
  ```

  大纲里用 `@chart: <类型>` + `@chart-data: <JSON>` 声明图表页，骨架自动落 chart-dominant 布局。
- **原生图标**：图标容器打空 `<svg data-icon="<集>:<名字>" viewBox=... width=.. height=..>`，
  `deck make --layouts` 后替换成 PowerPoint 原生 freeform 矢量图标（双击显示可编辑
  锚点，非图片）。内置 Phosphor（`ph:`，256 viewBox，填充）+ Lucide（`lu:`，24 viewBox，
  线形）双集（vendored 于 `src/offipy/assets/icons/`，更新见 `scripts/fetch_icons.py`）；
  图标颜色继承容器的 `color`（HTML 里 `color: var(--accent)` 即上主题色；容器未设色时
  缺省取当前主题的 `--accent`）。Lucide 线形图标按 round 线帽/拐角渲染（对齐源 SVG
  设计，非图片非贴图）。示例见
  [`examples/decks/icons/icons-demo.html`](examples/decks/icons/icons-demo.html)。

  ```html
  <svg class="icon" data-icon="ph:check-circle" viewBox="0 0 256 256" width="72" height="72"></svg>
  ```

  大纲里用 `@icons: <名字>[:标签]; ...` 声明图标行（默认 `ph:` 前缀），骨架自动落
  `icons-row` 布局（超过 3 个图标自动换行，3 列/行）。示例见
  [`examples/outline/icons-demo.md`](examples/outline/icons-demo.md)。
- **反馈学习**：审计后处置（fixed / accepted / ignored）记入 `~/.offipy/feedback.jsonl`，`feedback.dimension_weights()` 折算审计权重，越修越严（P2 验证版）

```python
from offipy import deck, aesthetic

pptx = deck.render("deck.html", theme="mckinsey", apply_layouts=True)  # 注入主题+布局
report = aesthetic.audit("deck.html")  # 读 measurement 打分
print(report.markdown())
```

## MCP server（Claude 接入）

`offipy mcp` 启动 MCP stdio server，把 Word / Excel / PowerPoint 全部操作暴露为 MCP 工具
（`ppt_set_title`、`word_write_line`、`excel_set_cell` 等）。工具调用与 `offipy` 命令等价，
作用在用户当前激活的文档 / 工作簿 / 演示文稿上，窗口实时可见。

在 Claude Desktop 的 `claude_desktop_config.json` 添加。`<OFFIPY_ROOT>` 是你本地仓库的绝对路径（Windows 示例：`C:\\path\\to\\offipy`），**不要提交真实本机路径**：

```json
{
  "mcpServers": {
    "offipy": {
      "command": "<OFFIPY_ROOT>\\.venv\\Scripts\\python.exe",
      "args": ["-m", "offipy.mcp_server"]
    }
  }
}
```

手动验证（无需 Office，仅握手）：

```bash
offipy mcp        # 阻塞运行，等待 stdio 客户端接入
```

## 发布硬化（ChatGPT 审核修复对照）

0.9.0a1 按第三方审核全量修复收口，逐项对应：

| 审核项 | 修复 |
|--------|------|
| P0-1 转换器未进 wheel | converter vendored 进 `src/offipy/_vendor/`，`pip install` 后 deck 管线可用 |
| P0-2 chromium 预检缺失 | 渲染前检查浏览器，缺失抛 `ConversionError` 并提示安装 |
| P0-3 库层抛 SystemExit | 新增 offipy 异常体系，库层不再 `sys.exit` |
| P0-4 server 安全 | Bearer token + `/status` + 16MB 上限 + Content-Type + op 白名单 |
| P0-5 mcp 依赖过宽 | `mcp>=2.0,<3.0` 收窄 |
| P0-6 发布无门禁 | release workflow 补全：lint→format→mypy→pytest→tag 版本校验→build→twine→安装冒烟→`--verify-tag` |
| P0-7 非 Windows 无法 import | 惰性 COM import，跨平台可用 |
| P1 全量 | CLI 改名 `offipy` + 复杂参数、高层 API、实时文档会话语义、save 覆盖保护、metadata/文档/类型/CI 治理 |
| P2 架构 | 会话语义、CLI 参数系统、高层 API 三项架构项落地 |

**round-3（ChatGPT_v3 审核）**：四条主线——目标身份 / 进程所有权 / 失败原子性 / 入口一致性，11 批按批收口：

| 批次 | 修复 |
|------|------|
| 进程所有权 + HTTP 边界 | `_probe` 四态（ok/auth_fail/mismatch/down）、`ensure_server` 不杀策略、鉴权 `/shutdown`、`server status` 只读、路径 404 / 负长度 400 / 响应上限 500 |
| 目标身份 + 只读不创建 | `get_target`、`/status` targets、`expected_target` 绑定、无文档抛 `TargetNotFoundError` |
| DisplayAlerts 作用域 + deck 原子 + CLI 真布尔 | op 内自存自还、临时文件 + `os.replace` 原子替换、`store_true` 布尔 |
| 线程 + worker + 返回 + 异常 | `ThreadingHTTPServer` + 单 COM worker 队列、`OperationResult` 契约、领域异常体系 A |
| schema 单一来源 | 新增 RPC 只改 `schema.py`，server/CLI/MCP 三入口派生 |
| extras + 品牌 | `office/deck/mcp/all` 拆分、`py.typed`、统一 `offipy` |
| CI 矩阵 + 非 Windows | pure-module（覆盖率门槛）/ windows 3.10–3.13 / wheel-smoke / office-real；MCP in-memory 协议层测试 |

## 开发

```bash
uv sync --extra dev                     # 装 dev 依赖（ruff / mypy / pytest）
uv run ruff check .                     # lint
uv run ruff format --check .            # 格式
uv run mypy src/offipy                  # 类型
uv run pytest tests -q                  # 测试（COM 集成测试无 Office 自动跳过）
```

贡献规范与门禁见 [`CONTRIBUTING.md`](CONTRIBUTING.md)。

## 结构

```
src/offipy/
  core.py       # COM 生命周期/会话管理 + active_doc/doc_alive 会话语义
  exceptions.py # offipy 异常体系（OffipyError + 10 子类，策略 A 领域异常）
  schema.py     # operation schema 单一来源（OpSpec 表，server/CLI/MCP 三入口派生）
  result.py     # OperationResult 返回契约（HTTP-only：ok/operation/resource_id/message/data）
  paths.py      # 用户数据目录 / converter 数据目录 / ensure_writable 覆盖保护
  server.py     # 常驻会话 HTTP server（token 鉴权 + worker 队列 + /status + /shutdown）
  cli.py        # `offipy` 命令入口（复杂参数：重复 flag → list / --payload JSON）
  api.py        # 高层 API facade：Excel() / Word() / Ppt() 上下文管理器
  mcp_server.py # MCP stdio server（三套件操作 → MCP 工具）
  excel.py / word.py / ppt.py   # 三套件原子操作
  client.py     # server 的 HTTP 客户端（HTML 管线复用）*
  envcheck.py   # `offipy check` 环境就绪诊断（分组清单 + --json）*
  deck.py       # HTML → 可编辑 PPTX 管线（render/open_live/export_slides）*
  design.py     # 设计系统：token 模型 + 3 套内置主题 + .slide 基础样式 *
  layouts.py    # 命名布局库：11 种布局组件 + data-layout 注入 *
  charts.py     # 原生图表：图表声明解析 + 原生可编辑图表注入（bar/line/pie）*
  icons.py      # 原生图标：SVG 路径压平 + freeform 矢量图标注入（ph/lu 双集）*
  assets/icons/ # vendored 图标资产（Phosphor fill + Lucide）+ manifest + LICENSE *
  aesthetic.py  # 审美审计：留白/字号层级/色数/对比度/一致性 → 打分报告 *
  autopick.py   # 自动选型：内容结构 → 推荐主题 + 每页布局 + 理由 *
  feedback.py   # 反馈学习：审计处置 → 维度权重（P2 验证版） *
  outline.py    # 内容工作流：markdown 大纲 → 逐页结构化内容 → HTML 骨架 *
  _vendor/html_to_editable_pptx/  # vendored HTML→PPTX 转换器（外协代码，MIT）*
tests/        # pytest
docs/         # 协议（protocol.md）/ 返回契约（api.md）/ 弃用（deprecation.md）/ 兼容矩阵（compatibility.md）/ 发布手册（release.md）
examples/     # 可运行示例（decks / outline / excel / word）
```

\* 由「HTML-first 可编辑 PPTX 管线」与「M1/M2/M3 内容工作流」计划新增。

## License

MIT

## 反馈与问题

- **Bug / 功能请求**：到 [GitHub Issues](https://github.com/Zn070515/office-kit/issues) 提交，
  附上 `offipy check --json` 输出与最小复现。
- **预发布版本**：TestPyPI 冒烟用 `scripts/pypi_smoke.py --index https://test.pypi.org/simple --version <预发布号>`
  （见 CHANGELOG「预发布编号策略」）。预发布版本仅供验证，稳定首发前 `__version__` / tag / CHANGELOG
  顶层三者保持一致。
