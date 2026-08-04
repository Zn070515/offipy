# PPT 功能生态调研

> 调研日期：2026-08-04
> 范围：offipy 现有 PPT 能力 vs 开源生态（python-pptx / COM 自动化 / HTML→PPTX / AI 生成），
> 依据：WebSearch + `gh search repos` 实跑数据（star 数以调研当日为准，会浮动）。

---

## 1. offipy 当前 PPT 能力盘点

### 1.1 会话式 COM 驱动（`src/offipy/ppt.py`）

在真实 PowerPoint 里以会话方式驱动当前文稿，跨进程通过 `ActivePresentation` 定位：

| 能力 | API |
|------|-----|
| 新建 / 打开文稿 | `new_pres()` / `open_pres(path)` |
| 定位当前文稿 | `active_pres()` |
| 保存（含另存路径） | `save(path=None)` |
| 导出 PDF | `save_pdf(path)`（`PP_SAVE_PDF=32`） |
| 逐页导出 PNG（供视觉迭代） | `export_slides(out_dir, width=1920, height=1080)` |
| 加幻灯片 | `add_slide(layout)`（1=标题 2=文本 5=仅标题 12=空白） |
| 标题 / 正文 / 备注 | `set_title` / `set_body` / `set_notes` |
| 文本框 / 图片 | `add_textbox` / `add_picture` |

核心价值：**操作的是活着的 PowerPoint**——不是离线的 OOXML 拼装，是用户亲眼看到页面在动的实况。输出直接就是原生 .pptx，可继续手改。

### 1.2 HTML→可编辑 PPTX 管线（`src/offipy/deck.py`）

Claude 写 16:9 HTML → vendored `third_party/html-to-editable-pptx`（即 Hasasasa/html-to-editable-pptx）转换 →
在真实 PowerPoint 打开实况 + 逐页导出 PNG 供视觉迭代。

- `render(html, out, only_slides, no_visual_audit, timeout)`：跑完整转换，返回 .pptx 绝对路径
- `make(html, out, open_live_flag, feedback_dir)`：render → 打开实况 → 可选导出 PNG 反馈
- 转换器是 **vector-first**：preflight → Playwright 实测 DOM 坐标 → 组装 → 内嵌字体 → 自检 → 视觉审计
- 产物为原生可编辑 .pptx（占位符、文本、图形都保留可编辑性），而非扁平化的图片页

> ⚠️ 已知边界：`third_party/` 转换器未进 wheel，`pip install offipy` 后 deck 管线不可用（PyPI 发布前置项，已延迟）。

---

## 2. 生态全景

PPT 自动化生态大致四条路线，按"离真实 PowerPoint 的距离"从远到近：

```
离线 OOXML 拼装（python-pptx）          —— 纯 Python 写 XML，无 Office
原生 API 绑定（Aspose.Slides 等）       —— 商业库，功能全但封闭
HTML/CSS → 转换（dom-to-pptx / html-to-editable-pptx）
                                        —— Web 技术栈出稿，vector-first 保可编辑
COM 驱动真实 Office（pywin32 / offipy）
                                        —— 操作活着实例，最接近"人用 PowerPoint"
AI 生成器（Anionex/banana-slides 等）   —— 上层封装，多数底层仍走 python-pptx/转换器
```

### 2.1 python-pptx 系（离线拼装主力）

- **scanny/python-pptx**（~3.5k★）：事实标准。纯 Python 读写作 .pptx（OOXML），可编辑性好，无 Office 依赖。
  - 优点：跨平台、可 CI、可批量、生态最成熟
  - 缺点：**不渲染**——不知道页面长什么样；排版全靠手算坐标；动画/字体测量不支持
- 上层的"skill 层"项目活跃度很高：
  - **GordenSun/GordenPPTSkill**（~2.9k★）：给 AI agent 用的 PPT 制作 skill，底层 python-pptx
  - **seulee26/mckinsey-pptx**（~540★）：麦肯锡风格的 PPT 模板/方法论，python-pptx 封装
  - **likaku/Mck-ppt-design-skill**（~240★）：同上赛道的 skill

> 判断：python-pptx 做"结构正确"很强，做"好看"很弱——因为没有渲染反馈闭环。offipy 的 deck 管线恰好补上这个闭环（Playwright 实测 + PNG 视觉迭代）。

### 2.2 HTML → PPTX 转换系

- **dom-to-pptx**（JS，Github 生态内较活跃）：浏览器 DOM → pptxgenjs，支持**动画/过渡**，但产物可编辑性一般
- **Hasasasa/html-to-editable-pptx**（offipy vendored 的那个）：vector-first，重点保**可编辑性**
- **Design-Arena/html-to-pptx**（~19★）：同类小项目
- WebSearch 未发现 Python 侧有超越 python-pptx + 自研转换的成熟现成方案——这条线基本是"各家用各自招"

> 判断：这条线的痛点就是可编辑性与美观的平衡。offipy vendored 方案选的是 vector-first + 自检 + 视觉审计，方向对，但**缺动画/过渡**（见 §4 差距）。

### 2.3 COM 驱动真实 Office

- **ykuwai/ppt-mcp**（~50★）：明确标注 "Real-time PowerPoint control via COM automation"——与 offipy 同路线
- **Ayushmaniar/powerpoint-mcp**（~110★）：COM/COM 类似方式驱动 PPT 的 MCP server

> 判断：COM 直驱真实 Office 的开源项目极少、star 普遍低，**但这是 offipy 与"人用 PowerPoint"对齐的唯一路线**。低 star 不代表没价值——说明这个方向做的人少，是差异化空间，也是为什么我们要自建会话 server（HTTP 127.0.0.1:8890 保活 COM 引用）。

### 2.4 AI 生成器

- **Anionex/banana-slides**（~15k★）：AI PPT 生成，web 前端 + 后端渲染
- **veasion/AiPPT**（~1.9k★）：AI PPT，中文友好
- **GordenSuperPPTSkills**（~1.7k★）：GordenPPTSkill 的增强版

> 判断：这个赛道拼的是"提示词→结构→模板库"的上层产品力，底层几乎都用 python-pptx 或自研转换。offipy 的定位不在这一层——我们是**底座**，让这些上层方案有"真实 PowerPoint 实况"这个可选项。

---

## 3. GitHub 头部项目（gh search 实跑数据）

> 数据源：`gh search repos`，star 数为调研当日值。仅列与 PPT 自动化直接相关的头部。

| 项目 | star | 路线 | 与 offipy 关系 |
|------|------|------|----------------|
| scanny/python-pptx | ~3.5k | 离线 OOXML | 生态基石，deck 管线的依赖之一 |
| GordenSun/GordenPPTSkill | ~2.9k | AI skill (pptx) | 竞品/参考：agent 出稿 |
| GongRzhe/Office-PowerPoint-MCP-Server | ~1.8k | MCP | **核心参考**：Office 全家桶 MCP，含 PPT |
| seulee26/mckinsey-pptx | ~540 | 模板方法论 | 设计参考 |
| likaku/Mck-ppt-design-skill | ~240 | skill (pptx) | 设计参考 |
| Ayushmaniar/powerpoint-mcp | ~110 | COM MCP | 同路线小项目 |
| ykuwai/ppt-mcp | ~50 | COM MCP | 同路线小项目 |
| Design-Arena/html-to-pptx | ~19 | HTML→PPTX | vendored 方案的同类 |

观察：
1. **MCP 是当前热度最高的整合形态**（Office-PowerPoint-MCP-Server ~1.8k★），offipy 的 COM 会话 server 天然可升级为 MCP server
2. **COM 直驱真实 Office 的专精项目极少且小**（ppt-mcp / powerpoint-mcp），offipy 是这个方向里工程化最完整的开源实现之一
3. **AI skill 层（2.9k★ / 1.7k★）说明"给 agent 用"是刚需**，但底层普遍离线拼装，无实况反馈

---

## 4. 差距分析

| 能力 | offipy | python-pptx | dom-to-pptx | AI 生成器 | 说明 |
|------|--------|-------------|-------------|-----------|------|
| 真实 PowerPoint 实况 | ✅ 核心 | ❌ | ❌ | ❌ | 唯一"页面在动"的路线 |
| 原生可编辑 .pptx | ✅ | ✅ | ⚠️ 一般 | ⚠️ | deck 管线 vector-first 保可编辑 |
| 渲染反馈闭环（视觉迭代） | ✅ Playwright + PNG | ❌ | ✅ 浏览器 | ✅ | offipy 出稿后能看、能改、能迭代 |
| 动画 / 过渡 | ❌ | ❌ | ✅ | ⚠️ | **最大功能差距** |
| 母版 / 模板复用 | ⚠️ 无专门 API | ✅ | ⚠️ | ✅ | 需补设计系统能力 |
| 逐页备注（演讲者备注） | ✅ set_notes | ✅ | ⚠️ | ⚠️ | 已有 |
| PDF 导出 | ✅ | ✅ | ⚠️ | ✅ | 已有 |
| MCP / agent 接入形态 | ⚠️ HTTP server 已备 | ❌ | ❌ | ✅ | 升级路径最短 |
| 跨平台（非 Windows） | ❌ | ✅ | ✅ | ✅ | 定位取舍，Windows-only |

### 结构性结论

1. **offipy 的护城河不是"能生成 pptx"，而是"生成后能在活 PowerPoint 里看到并继续改"**——这是所有离线方案给不了的。
2. **最大功能缺口是动画/过渡**：要追，得在 HTML 侧约定动画声明（如 `data-anim`）+ 转换器在 OOXML 里写 `<p:anim>` 序列；工作量大、收益待验证，建议留到设计体系成熟后。
3. **最值得立刻做的是 MCP 化**：`client.py` 已经是 HTTP 127.0.0.1:8890 的常驻 server，包一层 MCP 协议即可接入 Claude Desktop / Code 生态，直接对位 GongRzhe 那个 ~1.8k★ 项目。

---

## 5. offipy PPT 未来路线

按"投入产出比 + 差异化"排序：

### P0（低成本、高差异化）
- **MCP server 化**：把 8890 的 HTTP server 包成 MCP，让 Claude/其他 agent 通过标准协议驱动真实 PowerPoint。对齐生态最热形态，复用现有会话基础设施。
  - ✅ **已实现**（2026-08-04）：`office mcp` / `python -m offipy.mcp_server`，Word/Excel/PPT 全部操作暴露为 MCP 工具，见 [`../src/offipy/mcp_server.py`](../src/offipy/mcp_server.py)。
- **模板/母版 API**：`add_slide` 支持指定母版 + 自定义版式，把"设计系统"沉淀成可复用的 .potx。
- **逐页备注导出**（演讲者模式）配套：`export_slides` 时一并导出每页备注成 Markdown/PDF 讲稿。

### P1（中成本、高价值）
- **图片表格富内容**：`add_table`、图表（原生图表对象而非截图），让实况演示真正能承载数据页。
- **批量/脚本化**：`render` 管线支持批量 HTML → 多个 deck，服务 CI 出稿。

### P2（高成本、先验证后做）
- **动画/过渡**：HTML 侧 `data-anim` 声明 + 转换器写 `<p:anim>`。先做"淡入"一类高频动效，验证价值再铺开。
- **设计系统模板库**：对标 mckinsey-pptx 的思路，内置若干可主题化的 16:9 设计模板。

### 明确不做
- 跨平台：COM 是 Windows-only 的定位，跨平台会让会话式实况模型塌掉。
- 抢 AI 生成器的上层产品：我们做底座，不做 "输入主题出整稿" 的封装。

---

## 附：调研命令备忘

```bash
# 出站需走代理（本机 VPN 拦截出站）
export https_proxy=http://127.0.0.1:12334 http_proxy=http://127.0.0.1:12334

# 生态头部扫描
gh search repos "python-pptx" --sort stars --limit 10
gh search repos "powerpoint automation" --sort stars --limit 10
gh search repos "html to pptx" --sort stars --limit 10
gh search repos "ai ppt" --sort stars --limit 10
gh search repos "powerpoint mcp" --sort stars --limit 10
```
