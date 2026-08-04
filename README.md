# office-kit

Live Microsoft Office automation via COM（会话式驱动）。目标：让 Claude 能独立产出**美观、符合审美、言之有物**的 Office 产物（Word / PPT / Excel）。

> 库名 **offipy**（`pip install offipy`，`import offipy`）；CLI 命令 `office`。

> 当前状态：COM 会话管线（Excel / Word / PPT）已打通；「HTML-first 可编辑 PPTX」管线 + 设计系统已落地（背景见 [`docs/gap_analysis.md`](docs/gap_analysis.md)）。

## 特性

- **会话式常驻 server**：跨调用保持 Office 窗口存活、文档 / 工作簿 / 演示文稿状态不丢
- **三套件原子操作**：Word / Excel / PowerPoint 增删改 + 保存 / 导出 PDF
- **断连自愈**：用户关窗或 Office 退出后自动重建会话
- **HTML-first 管线 + 设计系统**：Claude 写 HTML 幻灯片 → 原生可编辑 `.pptx` → 实况展示 + 视觉迭代；内置设计 token、3 套主题、10 种布局、审美审计、自动选型、反馈学习（见下方「设计系统」）
- **MCP server**：把全部三套件操作暴露为 MCP 工具，Claude Desktop 等可直接驱动真实 Office

## 环境要求

- Windows + 已安装 Microsoft Office（Word / Excel / PowerPoint）
- Python ≥ 3.10（本仓库开发环境为 3.12）

## 安装

```bash
uv venv --python 3.12 .venv
uv pip install -e .
```

## 使用

```bash
# 首次调用自动在后台拉起常驻 server；之后所有操作打到同一进程
office excel new_book
office excel set_cell --sheet 1 --cell A1 --value 100
office excel format_cell --sheet 1 --cell A1 --bold true --size 14 --bg "#38BDF8"

office word new_doc
office word write_line --text "你好，世界"

office ppt new_pres
office ppt add_slide --layout 2
office ppt set_title --slide_idx 1 --text "标题"

office quit excel
```

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
- **命名布局库**：`hero-title` / `split-2col` / `cards-3` / `big-number` / `quote-frame` / `timeline` / `comparison` / `chart-dominant` / `portrait-feature` / `closer` —— `layouts.inject_layouts()` 按引用注入
- **审美审计**：转换产出的 measurement → 留白比例 / 字号层级数 / 每页色数 / 对比度 / 跨页一致性打分，`aesthetic.audit()` 输出带分报告供迭代
- **自动选型**：`autopick.pick()` 从内容结构推荐主题 + 每页布局 + 理由（纯规则，可覆盖）
- **内容工作流**：把「言之有物」的骨架标准化——写一份 markdown 大纲
  （`# 标题` + 每个 `## 段` = 一页，`- ` 要点、正文、`@layout:` 指令），
  `office deck outline --input outline.md --out deck.html` 一键得到 HTML 骨架，
  布局自动推断，再 `office deck make --theme <主题> --layouts` 注入主题与布局成稿。
  示例见 [`examples/outline/quarterly-review.md`](examples/outline/quarterly-review.md)。

  ```bash
  office deck outline --input examples/outline/quarterly-review.md --out out/quarterly.html
  office deck make --html out/quarterly.html --theme mckinsey --layouts --out out/quarterly.pptx
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
- **反馈学习**：审计后处置（fixed / accepted / ignored）记入 `~/.offipy/feedback.jsonl`，`feedback.dimension_weights()` 折算审计权重，越修越严（P2 验证版）

```python
from offipy import deck, aesthetic

pptx = deck.render("deck.html", theme="mckinsey", apply_layouts=True)  # 注入主题+布局
report = aesthetic.audit("deck.html")   # 读 measurement 打分
print(report.markdown())
```

## MCP server（Claude 接入）

`office mcp` 启动 MCP stdio server，把 Word / Excel / PowerPoint 全部操作暴露为 MCP 工具
（`ppt_set_title`、`word_write_line`、`excel_set_cell` 等）。工具调用与 `office` 命令等价，
作用在用户当前激活的文档 / 工作簿 / 演示文稿上，窗口实时可见。

在 Claude Desktop 的 `claude_desktop_config.json` 添加。`<OFFIPY_ROOT>` 是你本地仓库的绝对路径（Windows 示例：`C:\\path\\to\\officeforclaude`），**不要提交真实本机路径**：

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
office mcp        # 阻塞运行，等待 stdio 客户端接入
```

## 开发

```bash
uv sync --extra dev                     # 装 dev 依赖（ruff / mypy / pytest）
uv run ruff check .                     # lint
uv run ruff format --check .            # 格式
uv run mypy src/offipy                  # 类型
uv run pytest tests -q                  # 测试（COM 集成测试无 Office 自动跳过）
```

## 结构

```
src/offipy/
  core.py     # COM 应用生命周期与会话管理
  server.py   # 常驻会话 HTTP server（持有 COM 引用）
  cli.py      # `office` 命令入口
  mcp_server.py  # MCP stdio server（三套件操作 → MCP 工具）
  excel.py / word.py / ppt.py   # 三套件原子操作
  client.py   # server 的 HTTP 客户端（HTML 管线复用）*
  deck.py     # HTML → 可编辑 PPTX 管线（render/open_live/export_slides）*
  design.py   # 设计系统：token 模型 + 3 套内置主题 + .slide 基础样式 *
  layouts.py  # 命名布局库：10 种布局组件 + data-layout 注入 *
  aesthetic.py  # 审美审计：留白/字号层级/色数/对比度/一致性 → 打分报告 *
  autopick.py   # 自动选型：内容结构 → 推荐主题 + 每页布局 + 理由 *
  feedback.py   # 反馈学习：审计处置 → 维度权重（P2 验证版） *
  outline.py    # 内容工作流：markdown 大纲 → 逐页结构化内容 → HTML 骨架 *
  examples/outline/  # 内容工作流示例大纲 *
tests/        # pytest
docs/         # 差距分析与实施计划
third_party/  # vendored HTML→PPTX 转换器 *
```

\* 由「HTML-first 可编辑 PPTX 管线」与「M1 内容工作流」计划新增，见 [`docs/superpowers/plans/`](docs/superpowers/plans/)。

## License

MIT
