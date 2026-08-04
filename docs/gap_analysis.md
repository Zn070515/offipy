# OfficeForClaude — 差距分析与调研基线

> **目标**：让 Claude 独立产出「美观、符合审美、言之有物」的 Office 产物（Word / PPT / Excel）。
> **本文**：固定 2026-08-04 第一轮 websearch + gh 调研结论，作为后续开工的基线。第二轮同规格重跑的结果追加在「§3 第二轮验证」。

---

## 1. 项目现状（2026-08-04）

已打通**基础管线**（「能不能动」已解决）：

| 模块 | 作用 |
|------|------|
| `office_kit/` | COM 会话式自动化库（pywin32），Word / Excel / PPT 三套件 |
| `office_kit/server.py` | 常驻 session server（HTTP `127.0.0.1:8890`），持有 COM 引用跨调用保活、断连自动重建 |
| `office_kit/cli.py` | CLI 入口 `office <app> <op> [--key value ...]` |

已验证：三套件 new/open/save/save_pdf、用户关窗后自动重建会话、跨进程 Active 文档定位。

**目前能力**：基本增删改 + 保存导出 PDF。
**缺失**：无设计系统、无视觉反馈闭环、无图标 / 图表 / SmartArt / 动画 / 模板感知。

---

## 2. 差距在哪（第一轮调研结论）

### 2.1 三条成熟路线（2026 年现状）

| 路线 | 代表 | 核心思想 | 对「美」的答案 |
|------|------|---------|--------------|
| **HTML 画布设计** | Claude Design（Anthropic 官方，2026-04 发布） | 浏览器画布上用 HTML/CSS 做设计，接入设计系统（品牌色/字体/组件），对话内实时画布，导出 PPTX/PDF/HTML | **设计在 HTML** —— Claude 的设计能力长在 HTML/CSS，不长在 PowerPoint 绝对坐标上 |
| **HTML→可编辑 PPTX 管线** | `artifact-kit/html-to-pptx-skill` + `pptxgenjs-jsx` | 先写 HTML 幻灯片 → 浏览器测量真实 DOM 坐标（`data-ak-measure` / `readPptBox`）→ 用 pptxgenjs 重建为**原生可编辑** PPT 对象 | 不是截图管线；「HTML 设计 + 原生对象导出」，兼顾美感和可编辑 |
| **COM 精细操控** | `ykuwai/ppt-mcp`（156 工具 / 26 类） | 实时 COM 操控现有 PowerPoint，图标、SmartArt、动画、主题色感知 | 操控粒度取胜，覆盖超全 |

> 附：Gamma / Beautiful.ai / Tome 等 AI 演示工具同为「设计系统 + 模板约束」路线，本质还是模板库 + 自动布局，且 PPTX 导出保真度差。对自建管线参考价值低于上面三条。

### 2.2 差距五层

**① 设计资产层（最大差距）**
- 现在让 Claude 用绝对坐标摆 textbox = 让设计师拿尺子和打字机干活。
- 前端共识：**Claude 的设计能力在 HTML/CSS，因此「设计在 HTML → 转换到 Office」是「美」的答案**（Claude Design、artifact-kit、OpenDesign 全部收敛到这条）。
- 无设计令牌系统（色板 / 字号阶梯 / 间距网格 / 版式母版）。
- 无模板/母版感知：ppt-mcp 有 `ppt_activate_presentation` 锁定文件 + 读取主题色（`accent1`/`accent2`），我们只有硬编码 RGB。

**② 转换/渲染层（完全没有）**
- HTML → 可编辑 PPTX（测量 DOM + pptxgenjs）
- HTML → PDF（Chrome print CSS，适合报告/长文档）
- Puppeteer 逐页截图 → 图放进 PPTX（Claude Design 对复杂布局的实际做法；`bluzir/claude-code-design` 有开源复刻）
- 现状：我们只有 COM 原生 `save_pdf`，生成侧没有任何 HTML 渲染路径。

**③ 能力层（COM server vs ppt-mcp 的量化差距）**

| 能力 | 我们 | ppt-mcp |
|------|------|---------|
| 操作总数 | ~15 个基本操作 | 156 工具 / 26 类 |
| 图标 | 无 | Material Symbols 2500+ 图标 SVG |
| 主题色 | 无（硬编码 RGB） | 主题色感知 |
| 图表 / 表格合并 / 对齐分布 | 无 / 无 / 无 | 原生图表、批量表格、合并单元格、对齐 |
| SmartArt / 动画 / Freeform | 无 | 全套 |
| Excel 侧 | 仅 `format_cell` 基础格式 | —（我们缺图表、合并、边框、条件格式、冻结窗格、打印） |
| Word 侧 | 仅样式编号（Heading1-3） | —（缺封面、目录、页眉页脚、样式系统、主题） |

**④ 验证/反馈闭环（没有闭环就没有「美」）**
- Claude 现在看不到输出，无法迭代「美不美」。
- Claude Design 的实时画布就是这个闭环；其 PPTX 导出复杂布局会漂移，需「导出后人工/自动化终检」。
- 我们的 COM server 的独特价值：**打开真实 Office → 截图 → 喂回给 Claude → 它自己看自己迭代**——真实渲染验证，截图管线替代不了。

**⑤ 集成层**
- 现在 CLI `office word add_heading ...` 对 Claude 不友好。
- 标准做法是 MCP server（结构化工具 schema），Claude Code / Claude Desktop 原生驱动。ppt-mcp 即此形态。

### 2.3 战略判断

**「美」的答案不是把 COM 摆坐标练到极致，而是把设计引擎换成 HTML-first，让 COM 转职。**

- **生成侧**：Claude 写 HTML/CSS 幻灯片（设计令牌驱动）→ 转换 PPTX/PDF。
- **COM 侧**（不白做）：生成后打开真实 Office → 截图验证 → Claude 迭代微调（字号/对齐/配色）→ 导出 PDF。守住「看到页面在动」的诉求。
- **内容侧**：think-first, design-second —— 内容在对话里写 → 转 markdown outline（一节一页）→ 再进设计。

### 2.4 建议演进路径

1. **内容工作流**：outline → 逐页内容（这是「言之有物」的骨架）。
2. **设计引擎**：HTML-first + 设计令牌系统 → 转 PPTX（可编辑管线 vs 截图管线，先选一条打通）。
3. **验证闭环**：COM server 升级为「打开 → 截图 → 喂回 → 迭代」。
4. **能力补齐**：给 COM server 加图标 / 主题色 / 图表（对标 ppt-mcp 精选子集）。
5. **集成**：MCP 包装，Claude Code 原生驱动。

> **推荐第一步**：先做 ②「HTML-first 设计引擎」最小闭环（Claude 写 HTML 是真本事，出效果最快），再决定可编辑 vs 截图管线。

---

## 3. 第二轮验证（同规格重跑，2026-08-04）

**结论一：关键结论稳定复现。** Claude Design / Claude for PowerPoint / Gamma·Beautiful.ai·Tome 与第一轮一致；gh `powerpoint mcp` 头部仓库排序稳定（Ayushmaniar 107★、Ichigo3766 53★、ykuwai 50★）。两个空 gh 查询（`office automation claude` / `slide generation llm`）二次仍为空——判定为**真无结果**，非查询抖动。

**结论二：第一轮 4 个空 Web 查询全部补上，且炸出一个成熟工具簇**——「HTML → 可编辑 PPTX」不是少数人实验，而是**成熟开源生态**：

| 工具 | 方式 | 可编辑？ | 要点 |
|------|------|---------|------|
| **dom-to-pptx** | 客户端 DOM→原生形状（"Coordinate Scraper & Style Engine"） | ✅ 全可编辑 | 230k+/月下载；渐变/阴影/圆角/响应式；20+ 动画、70+ 转场；CLI `npx dom-to-pptx export slides.html`；自带 Claude Code agent-skill 安装器 |
| **html-to-editable-pptx** | Python CLI + Claude Code skill，vector-first | ✅ 文字/几何原生，复杂 CSS 效果局部快照垫底 | **与我们栈最对口（纯 Python）**；字体子集化；HTML/PPT 并排视觉审计管线 |
| **llm-dom-to-pptx** | 专为 LLM 生成的 HTML 设计 | ✅ | 自带 system prompt |
| **@0-ai/slide-gen** | 无头 Chromium 渲染→重建原生对象 | ✅ | Bun CLI；flex/grid/渐变/Google Fonts |
| **deckforge** | Chromium/CDP 三阶段管线 | ✅ | 128 个内置 Google 字体族 |
| **html2pptx-pro** | DOM 克隆+计算样式→pptxgenjs | ✅ | 100+ CSS 属性 |
| **SlideViber** | SVG 作稿（~10× 省 token）→ 原生 OOXML | ✅ | 21 套内置主题 |
| **html2ppt** | HTML→Chrome 截图→python-pptx | ❌ 图片 | 36+ 开源风格模板，100% 保真 |

**结论三：设计系统 ≠ 只是颜色。** huashu-slides 强调设计系统要含视觉哲学、字体比例、构图规则、情绪意图（18 套完整风格 DNA：Swiss 网格 / Ligne Claire / Fathom 数据叙事 / 极简奢侈风…）。

**结论四：可量化审美规范已存在**（pptx-generation skill 等）——正好是我们缺的「符合审美」的可验收标准：
- 对比度 ≥ 4.5:1；正文 ≥18pt、标题 ≥24pt；边距 0.5"；每页要点 ≤3 条
- 生成前先写设计计划（布局哲学 / 精确色板 / 字号层级 / 视觉强调策略）
- **validation gates**：对比度检查、字号校验、边距测量、元素重叠检测

**结论五：Marp / Reveal.js / Slidev 确认 —— PPTX 导出全是「逐页图片」，不可编辑。** 佐证原生可编辑必须走 pptxgenjs / python-pptx 级重建（OpenDesign 判断依旧成立）。Marp 导出最全（HTML/PDF/PPTX/PNG）。

### 对演进路径的修正
可编辑管线无需从零造：`dom-to-pptx`、`html-to-editable-pptx`（Python + Claude Code skill）可直接借鉴或包装。**推荐下一步先解剖 `html-to-editable-pptx`** —— 纯 Python、vector-first、带视觉审计，与我们的 COM server 互补（它生成，我们验证/微调/导出）。

---

## 4. Sources（第一轮）

- [Introducing Claude Design by Anthropic Labs](https://www.anthropic.com/news/claude-design-anthropic-labs)
- [Claude for PowerPoint vs Microsoft 365 Copilot (2026)](https://theaiagentindex.com/compare/claude-for-powerpoint-vs-microsoft-365-copilot)
- [Claude Design for Presentations: A Two-Tool Workflow](https://verygood.ventures/blog/building-presentations-with-claude-design/)
- [artifact-kit/html-to-pptx-skill](https://github.com/artifact-kit/html-to-pptx-skill)
- [artifact-kit/pptxgenjs-jsx](https://github.com/artifact-kit/pptxgenjs-jsx)
- [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- [ykuwai/ppt-mcp](https://github.com/ykuwai/ppt-mcp)
- [Pandemonium-Research/OpenDesign](https://github.com/Pandemonium-Research/OpenDesign)
- [bluzir/claude-code-design](https://github.com/bluzir/claude-code-design)
- [Beautiful.ai vs. Gamma](https://mktg.beautiful.ai/comparison/beautiful-ai-vs-gamma)
- [Best AI Presentation Tools in 2026](https://awesomeagents.ai/tools/best-ai-presentation-tools-2026/)

### Sources（第二轮）
- [dom-to-pptx (GitHub)](https://github.com/atharva9167j/dom-to-pptx)
- [html-to-editable-pptx (GitHub)](https://github.com/Hasasasa/html-to-editable-pptx)
- [html2ppt (GitHub)](https://github.com/AnonymousStudy972/html2ppt)
- [SlideViber skill (GitHub)](https://github.com/tf71991/slideviber-skill)
- [guizang-ppt-skill (GitHub)](https://github.com/SeanDongX/guizang-ppt-skill)
- [llm-dom-to-pptx (GitHub)](https://github.com/wenson0106/llm-dom-to-pptx)
- [Marp vs Reveal.js (classmethod)](https://dev.classmethod.jp/en/articles/marp-vs-revealjs-markdown-presentation-tools-comparison/)
- [Slidev vs Marp vs Reveal.js 2026 (pkgpulse)](https://www.pkgpulse.com/guides/slidev-vs-marp-vs-revealjs-code-first-presentations-2026)
- [PptxGenJS (npm)](https://www.npmjs.com/package/pptxgenjs)
