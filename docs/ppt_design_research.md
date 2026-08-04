# PPT 艺术层（设计/审美）生态调研

> 调研日期：2026-08-04
> 范围：PPT 的**视觉设计层面**——排版/网格/色彩/留白方法论、可编码的风格规范（麦肯锡等）、
> AI 生成 PPT 的审美提升趋势、16:9 设计系统/tokens。
> 依据：WebSearch + `gh search`/`gh repo view` 实跑数据（star 数以调研当日为准）。
> 上一篇技术层调研见 [`ppt_research.md`](ppt_research.md)。

---

## 1. offipy 现状（艺术层面盘点）

offipy 的 PPT 产物有两条路，艺术能力都还处于「刚够用」：

| 能力 | 现状 | 缺口 |
|------|------|------|
| HTML-first 转换（`deck.py`） | HTML/CSS 由 Claude 直接写，天然表达设计意图 | 无内置设计 token / 风格主题，每次从零排版 |
| 视觉审计（`--no-visual-audit` 反向） | 转换器有 self_check + visual_audit | 审计偏「技术正确」（元素是否落位），非「审美得分」 |
| 逐页导出 PNG（`export_slides`） | 1920×1080 实况页导出，供视觉迭代 | 已有——这是审美提升的**基础设施** |
| 会话式实况展示 | 真实 PowerPoint 里看效果 | 独有优势，AI 出稿后能人眼复核 |

关键结论：offipy 是 **HTML-first**，而「设计系统」在 HTML/CSS 里就是天然载体（token = CSS 变量，布局 = 组件）。艺术层路线基本是「往 deck 管线里加设计系统」，不用另起炉灶。

---

## 2. 设计方法论共识（可执行铁律）

多个独立来源（设计博客、学术、AI 出稿实践）高度一致，可提炼成**编码级规则**：

### 2.1 一致性与四要素
专业感来自「一次决定、处处贯彻」：**spacing（间距）、type hierarchy（字号层级）、alignment（对齐）、restraint（克制）**。间距不一致是「最响亮的业余信号」。

### 2.2 一页一观点（Billboard Test）
- 每页 ≤3–5 秒能读出核心观点（Steve Jobs 3 秒规则）
- 每页一个强标题，直接给出结论（「68% 客户准备切换」而非「客户行为分析」）
- 内容块 ≤7±2（理想 3–5）；**留白 ≥40%**；5% 边缘安全区

### 2.3 排版
- 最多两种字体：标题 + 正文（衬线标题 + 无衬线正文是咨询风标准组合）
- 用字重建层级，不靠加字体
- 模块化字号（1.25–1.618）；一页 ≤4 种字号
- 标题 ≥48px、正文 ≥24px（16:9 全屏）；行高 1.4–1.6；行长 ≤60 字符
- 标题 36–44pt / 正文 20–24pt（传统 PowerPoint 习惯），投影场景最小 14pt

### 2.4 色彩
- 60-30-10 分配（主色 60 / 辅色 30 / 强调 10）；一页最多一个强调色
- 全 deck 2–3 色：主品牌色 + 强调色 + 中性（黑/炭/浅灰）
- 强调色只给关键数据/结论，其余保持中性（「highlight the one」）
- 正文对比度 ≥4.5:1（WCAG AA，目标 7:1 AAA）；**绝不用色相单独编码信息**（色盲友好）
- 亮底（白/米/浅灰）最稳；深底（藏青/炭/黑）配浅色字，出彩但挑环境

### 2.5 网格与留白
- **8pt 网格**：所有间距对齐到 8pt；相关元素 ≤16px 距、无关元素 ≥48px 距
- 左对齐优先（居中是最后手段）；F-pattern：标题 + 主视觉在左上
- 数据墨水比 ≥80%（去掉边框/网格线/冗余图例）
- 一页一观点，塞不下就拆两页

### 2.6 数据页
- 标题即洞察；柱状图默认（「拿不准就用柱状图」）；直标数据、标来源

---

## 3. 可编码风格：麦肯锡（最成熟的参考）

麦肯锡 2019 年视觉体系（Wolff Olins 设计）+ 公开模板还原值，可直接作为 offipy 第一套内置主题：

### 色彩 token

| 角色 | Hex | 用法 |
|------|-----|------|
| 深藏青 hero/章节背景 | `#051C2C` | 章节页/封面，与白底混搭制造节奏 |
| 电光蓝 强调/CTA | `#2251FF` | **克制使用**，只高亮关键数据 |
| 深底图表背景 | `#0C1C23` | 深色页图表 |
| 白 主背景 | `#FFFFFF` | 亮底主力 |
| 正文 近黑 | `#222222` | 文本几乎只用黑 |
| 次要文本 中灰 | `#A2AAAD` | 注释/来源 |

学术色相分析：主色调蓝紫系（PB），亮度高位集中，**饱和度普遍低（0–3）**——克制是风格核心。

### 字体
- 标题：Bower（衬线，替代 Georgia）；正文：McKinsey Sans（替代 Arial）
- 全页字号统一（常见 14pt），不用大字强调、靠结构

### 结构规则
- **行动导向标题**（「建议探索 A/B 两个服务」而非「有四种服务值得探索」）
- 金字塔原理：结论先行 → 关键点 → 数据支撑
- 一图一意一格式；柱状图默认；图标极简单色、尺寸统一
- 白/深蓝页**混排**制造章节对比；柔和渐变增强层次
- 动画克制，仅助理解才用

---

## 4. AI 生成 PPT 的审美提升（2026 趋势）

### 4.1 「让 AI 做的看起来不像 AI 做的」
- **ItsssssJack/power-design**（587★）把规律固化：**Brand DNA × 20 条设计原则**（源自 Tufte、Duarte、Reynolds、Müller-Brockmann、WCAG 2.2）——一致性、60-30-10、8pt 网格、≥40% 留白、3 秒可读全部量化成可执行规则
- 共识：AI 出稿的审美短板不是「不会画」，是**不守纪律**——规则一旦被系统性锁住，观感就上去

### 4.2 学术：审美可量化、可学习
- **AeSlides**：用 RL + verifiable reward 量化布局质量，低成本捕获版式缺陷
- **EvoPresent（ICLR 2026）**：审美感知 RL 模型 PresAesth 打分 + 缺陷修正 + 对比反馈；发现**高质量反馈是自我提升的关键**
- **DeepSlides（ACL 2026）**：模板无关分层生成，多 agent RL 提升审美
- 启示：offipy 的 `visual_audit` + 逐页 PNG 已经是「反馈信号」，下一步是把它升级成**可打分的审美审计**

### 4.3 出稿工作流共识
Brief（受众/目标/语气）→ 先定结构 → 重写标题（给结论）→ 统一视觉系统 → 换掉凑数图片 → 数据页做减法 → 朗读排练、细节进备注。这跟 offipy「HTML 写稿 → 实况展示 → PNG 迭代」天然同构。

---

## 5. 16:9 设计系统 / tokens（实现模式）

### 5.1 设计 token 层（先定 token 再排页面）
- 色彩角色：background / primary / accent / body / muted / divider / highlight
- 字号刻度：title / section / body / label / citation（如 26 / 22 / 20 / 16 / 13pt）
- 间距：8pt 基数、页边距常量（如 0.5"）
- 圆角/内边距 token（minimax-ai 风格：Sharp & Compact / Soft & Balanced / Rounded & Spacious / Pill & Airy 四模式）

### 5.2 命名布局库（可组合的页面模式池）
agentara/skills 给出了 10 种可复用布局：`hero-title`、`split-2col`、`3-card`、`big-number`、`chart-dominant`、`quote-frame`、`timeline`、`portrait-feature`、`comparison`、`closer`——每种带解剖图 + do/don't + 构图启发式。

### 5.3 实现载体
- HTML/CSS（offipy deck 管线的天然载体）：token = CSS 变量，布局 = 组件
- PptxGenJS / PowerPoint Slide Master / Typst / Figma Variables 是其他生态的等效物
- 画布统一 16:9：1920×1080 px 或 10" × 5.625"

---

## 6. GitHub 头部项目（gh 实跑数据）

| 项目 | star | 方向 | 对 offipy 的启示 |
|------|------|------|------------------|
| ItsssssJack/power-design | 587 | Claude skill：Brand DNA × 20 条设计规则 | **直接参考**：规则量化成可执行清单 |
| seulee26/mckinsey-pptx | 544 | 40 个麦肯锡模板 + 选模板 subagent | 模板库 + 自动选型的设计 |
| agentara/skills | 427 | 通用 AI builder skills（含 presentation-design） | 10 种命名布局库可借鉴 |
| seacen/deckify | 52 | URL → HTML slide 设计系统 | 同为 HTML-first 设计系统，对标 |
| SlideSpeak/slide-design-skill | 新（高活跃） | describe → bespoke style → 1920×1080 HTML slides | 风格推导 + 真实图表渲染 |
| COREkin/mckinsey-pptx-mcp | 0 | MCP server：22 模板 | 模板化 + MCP 的组合被验证 |
| edu-ai-builders/visual-cognition-slides | 新 | 认知科学/教学设计 HTML slides | 教学设计维度 |
| Rikuto-des/slide-mcp | 新 | Figma Slides via MCP，31 布局 | 布局库 + MCP |
| Dhruv2541/AI-PPT-Template-Designer | 新 | AI 模板设计器 | 模板生成 |

观察：
1. **规则量化是护城河**：power-design 587★ 证明「把设计原则编码成规则」是被认可的方向
2. **模板 + 自动选型**是 mckinsey-pptx 的模式（subagent 选模板并解释），offipy 可借鉴
3. **HTML-first 设计系统**是小而新的赛道（deckify 52★），offipy 正好在这条线上

---

## 7. 差距分析

| 能力 | offipy | power-design | mckinsey-pptx | 说明 |
|------|--------|--------------|---------------|------|
| 视觉迭代闭环（看效果） | ✅ 实况 + PNG | ❌ | ❌ | offipy 独有 |
| 可编码设计规则（token/20条） | ❌ 无内置 | ✅ | ⚠️ 模板内置 | **最大差距** |
| 风格主题库（咨询/学术/极简） | ❌ | ⚠️ Brand DNA | ✅ 麦肯锡 | 需内置 2–3 套 |
| 命名布局库 | ❌ | ⚠️ | ✅ 40 模板 | 10 种布局起步 |
| 审美自检（打分式） | ⚠️ visual_audit 偏技术 | ⚠️ | ❌ | 升级空间大 |
| 自动选型（主题/布局） | ❌ | ❌ | ✅ subagent | 中期目标 |
| 模板文件复用（.potx） | ❌ | ❌ | ⚠️ | 需补 |

### 结构性结论
1. **offipy 的独有优势是「看完再改」的闭环**——AI 出稿、实时可见、逐页 PNG 复核。这是 power-design / mckinsey-pptx 都没有的。艺术层的提升应该**挂在这个闭环上**：先出稿，再按规则审计，再迭代。
2. **最大缺口是「无内置设计系统」**：没有 token、没有主题、没有布局库，Claude 每次从空白排版，一致性靠运气。而 HTML-first 让 token/布局库的实现成本极低。
3. **最值得抄的是 power-design 的「规则量化」思路**：把第 2 节的铁律转成 offipy 自己的 `visual_audit` 规则，让审美有客观评分。

---

## 8. offipy 艺术层路线

### P0（低成本、直接提升出稿质量）
- **内置设计 token 层**：`deck` 管线注入默认 CSS 变量集（色彩角色 + 字号刻度 + 8pt 间距 + 16:9 画布），Claude 写 HTML 时直接引用，天然一致。
- **内置 2–3 套风格主题**：麦肯锡蓝（第 3 节 hex/字体）+ 学术极简 + 深色科技，主题 = 一组 token 覆盖。
- **把 20 条设计规则转成可审计项**：升级 `visual_audit` 从「技术检查」到「审美打分」（留白比例、字号层级数、每页色数、对比度），输出带分数的审计报告供迭代。

### P1（中成本、高价值）
- **命名布局库**：内置 10 种布局组件（hero-title / split-2col / 3-card / big-number / timeline / comparison…），Claude 通过 `data-layout` 引用。
- **视觉一致性校验**：转换后自动比对全 deck 的 token 一致性（字号/颜色/间距是否漂移），漂移项列入审计报告。

### P2（高成本、先验证后做）
- **自动选型**：像 mckinsey-pptx 的 subagent——根据内容结构/受众自动选主题 + 布局并解释理由。
- **审美 RL 自学习**（对标 AeSlides/EvoPresent）：用 PNG 视觉反馈训练打分器，让审计越来越接近人眼。

### 明确不做
- 不搞「AI 一键生成整稿」的上层产品（我们是底座）。
- 不做模板文件（.potx）优先路线——HTML token 方案维护成本更低、与现有管线同构。

---

## 附：调研命令备忘

```bash
# 出站需走代理（本机 VPN 拦截出站）
export https_proxy=http://127.0.0.1:12334 http_proxy=http://127.0.0.1:12334

# 生态扫描
gh search repos "slide design skill" --sort stars --limit 8
gh search repos "mckinsey pptx" --sort stars --limit 8
gh search repos "slide design system" --sort stars --limit 8
gh search repos "ppt template design" --sort stars --limit 8
gh repo view ItsssssJack/power-design --json stargazerCount,description
```
