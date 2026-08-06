# 艺术分析（offipy.art）

`offipy.art` 是一套**纯标准库、确定性、只建议**的视觉/排版质量分析：不调用 AI、不依赖
Microsoft Office、`import offipy` 不加载 python-pptx。它把一页幻灯片抽象成「场景
（ArtScene）」，用确定性规则评估 5 个维度（层级 / 构图 / 排版 / 颜色 / 媒体），产出
分页报告。

- **只建议不阻断**：每条 finding 带 `confidence` 与 `severity`，但 art 层**不提供总分门禁**，
  不做「及格 / 不及格」判定——取舍留给调用方（如 deck 管线把它当生成后的质量参考）。
- **确定性**：同一输入必得同一输出，规则无随机、无模型、无网络。
- **证据诚实**：证据不足的维度降级为 `insufficient_evidence`，绝不靠猜补数误报。

## 安装与依赖

art 层**零额外依赖**（纯 stdlib），`import offipy` 即可用：

```python
import offipy
offipy.analyze_scene   # 存在即证明 art 已就绪
```

两个证据源（见下）按需选装：

- `measurements`（HTML→PPTX 管线的 DOM 测量 JSON）——`deck` 管线自带；
- `pptx`（PPTX 几何审计报告）——解析 `.pptx` 需要 python-pptx（`pip install "offipy[deck]"`），
  但只有真正解析文件时才加载。

## 快速开始

```python
from offipy import build_scene, analyze_scene, render_markdown

# 1) 建场景：测量数据（浏览器渲染的真实像素证据）为主、几何审计为辅
scene = build_scene(
    measurements="out/report_audit/_cache/measurements.json",  # deck 管线落盘位置
    pptx="out/report.pptx",
)
# 2) 分析
report = analyze_scene(scene, profile="balanced")
# 3) 报告
print(render_markdown(report))
```

组合入口（一次调用同时做几何审计 + 艺术分析）：

```python
from offipy import analyze_deck

report = analyze_deck(
    pptx="out/report.pptx",
    measurements="out/report_audit/_cache/measurements.json",
    profile="consulting",
)
print(report.art.slides[0].by_dimension("color").status)   # assessed / insufficient_evidence
for s in report.art.slides:
    for d in s.dimensions:
        for f in d.findings:
            print(s.slide_index, d.dimension, f.rule_id, f.confidence)
```

只给 `pptx=`（无测量数据）时，依赖字号/颜色证据的维度自动降级
（`insufficient_evidence` + `art.evidence.limited` warning），纯几何规则照常运行：

```python
report = analyze_deck(pptx="external.pptx", profile="balanced")
```

## 证据源与坐标约定

ArtScene 由两类证据源构建，按「测量为主、审计为副」合并：

| 源 | 证据 | 单位 | 说明 |
|----|------|------|------|
| `measurements`（MeasurementAdapter） | 颜色、字号、自然尺寸、文本、字体族 | px（归一化到 [0,1] 分数） | HTML→PPTX 管线浏览器渲染测量的真实像素证据；role 词表 title/body/subtitle/image/shape |
| `pptx`（PptxAuditAdapter） | 几何（位置/尺寸/role 分类）、文本 | pt（归一化到 [0,1] 分数） | `audit_pptx` 的几何快照；无字号/颜色证据；role 词表 background/header/footer/page_number/title/content/decoration/unknown |

合并（`merge_scenes`）按**文本强佐证 + 几何兜底**做一对一匹配：归一 role 相同且文本相同
→ 身份 0.8；跨词表文本相同 → 0.7；双方都有文本但不同 → 不合并；至少一边无文本 →
几何邻近（中心距 ≤0.2）+ role 加分。未匹配的审计元素保留并附 `art.merge.unmatched`
warning——绝不静默丢弃。

坐标统一归一化为幻灯片宽高分数（[0,1]），`font_size_norm = font_size / 页高`（无量纲），
规则只比归一值、不跨源直接比原始尺寸。

## grade / confidence / evidence_coverage 三分离

每维度（DimensionAssessment）与每条 finding 的语义严格分离，三者互不冒充：

- **grade**：`excellent / good / attention / poor`——只表**质量**。由该维度 assessed 规则的
  严重度加权分累计（LOW=0.5 / MID=1.5 / HIGH=3.0），阈值 `(0, 1.0, 2.5)` 分段。
- **confidence**：`(0, 1]`——只表**可信度**。证据覆盖越高、越贴合阈值边缘的 finding 越可信；
  实验性规则强制 `conf ≤ 0.3`。
- **evidence_coverage**：`[0, 1]`——只表**证据覆盖**。由规则自己上报 eligible/covered，
  不猜常数；**coverage < 0.5 → 维度降级 `insufficient_evidence`，该维度 findings 丢弃**。

`experimental_score`（0-100）仅在 ≥3 个维度 assessed 时返回，只用于排序/对比，**不对外宣称
客观美学分数**。

## 维度与规则

| 维度 | 规则 | 说明 |
|------|------|------|
| hierarchy | no_focus / focus_conflict / active_title | 焦点是否明确、是否与标题冲突 |
| composition | off_balance / corner_cluster / spacing_drift | 重心、角落聚集、间距漂移 |
| typography | tiny_text / many_families / flat_scale | 字号过小、字体族过多、层级扁平 |
| color | no_accent / accent_flood / low_contrast | 无强调色、强调色过载、前景/背景对比不足 |
| media | distorted_image / oversized_image / image_overlap | 图片失真、过大、互相重叠 |

规则全部**确定性**：同一场景必得同一结果。全部 `rule_id` 冻结（含 5 条 experimental，
`conf ≤ 0.3`、不驱动任何降级判定）。

## 内置 profile

`offipy.profile_names()` → `['balanced', 'consulting', 'academic', 'technology', 'event']`。
profile 决定规则启用集、实验性规则与阈值，`get_profile(name)` 可读、可扩展。

## Deck 一致性

`analyze_scene` 在分页规则之外还评估跨页一致性（`assess_deck`）：按 role 分组（≥3 个元素）
检查间距 / 字号 / 背景色漂移，产出 `report.deck_findings`。

## 与 deck 管线集成

`deck.render_with_quality_report(html, out=..., audit_mode=..., fail_on=..., profile=...)`：
HTML→PPTX 生成后**同时**产出几何审计（`audit_pptx`）与艺术分析（`build_scene` +
`analyze_scene`），返回 `QualityRenderResult`（含 `art_report` / `deck_quality`）——
生成即质量参考。`audit_mode="report"` 只报告、`"strict"` 达 `fail_on` 门槛才退出非 0。

## 已知边界

- **PptxAuditAdapter 无字号/颜色证据**：只传 `pptx=` 时，hierarchy / typography / color
  维度证据不足 → `insufficient_evidence`，不是误报；要完整评估请给 `measurements=`。
- **RenderedSlide（PNG/slides_dir 像素级）分析延后到 v0.12.1**：`build_scene(slides_dir=...)`
  明确拒绝（`InvalidArgumentError`）。
- **双源合并的未匹配元素**：测量未建模的审计装饰元素（如装饰线）保留 + warning，
  不静默丢弃，也不稀释已匹配元素的证据。
