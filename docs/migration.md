## v0.11.x → v0.12.1

- `ART_SCHEMA_VERSION` / `ART_REPORT_SCHEMA_VERSION` → `"0.2"`：0.1 报告仍可被 0.2 读取
  （像素证据字段默认 None）；compare 跨 schema 给出 `art.compare.schema_mismatch` 建议性 warning。
- `build_scene` / `analyze_deck` 新增 `slides_dir=`：读逐页 PNG 像素证据（需要 Pillow，`offipy[deck]`）。
- `render_with_quality_report` 新增 `pixel_analysis="off"|"best_effort"|"required"`、
  `preserve_pixel_slides`、`slides_output_dir`：默认 `off`，行为不变。
- `DimensionAssessment` 新增可选 `reliability` / `minimum_reliability`；`ArtFinding` 新增
  `evidence_sources` / `evidence_reliability` / `evidence_method`（仅像素路径非空）。

---

# 0.11 → 0.12 迁移指南

0.12 是**纯新增**版本：不破坏任何 0.11 的既有 API、CLI 行为或返回契约。现有代码**无需改动**。

## 新增（全部可选项）

- **offipy.art 艺术分析**：`build_scene(measurements=..., pptx=...)` 建场景 +
  `analyze_scene(scene, profile=...)` 评估（5 维度规则），grade / confidence / evidence_coverage
  三分离、证据不足降级 `insufficient_evidence`；**只建议不阻断**，无总分门禁。纯标准库，
  `import offipy` 即有（不加载 python-pptx / AI / COM）。
- **组合入口**：`analyze_deck(pptx=..., measurements=..., profile=...)` 一次调用几何审计 + 艺术分析，
  产出 `DeckQualityReport`（`.geometry` / `.art` / `.warnings`）。
- **生成即质量参考**：`deck.render_with_quality_report(html, audit_mode=..., fail_on=..., profile=...)`
  ——**并行新增入口，`render_with_report` 契约完全不变**。`render_with_quality_report` 在几何审计之外
  再产出艺术分析，返回 `QualityRenderResult`（含 `art_report` / `deck_quality`）。
- **内置 profile**：`balanced` / `consulting` / `academic` / `technology` / `event`。
- **基线对比 v2**：`compare_reports(before, after)` 产出 `ArtReportDiff`。
- CLI / RPC / MCP 三入口行为不受影响；完整 art API 见 [docs/art.md](art.md)。

## 迁移步骤

0.11 → 0.12 **无迁移步骤**。如果 0.11 代码能跑，0.12 直接换版本号即可。

唯一注意点（非破坏）：

- 想给 deck 生成附带艺术分析，把 `render_with_report` 换成新的
  `render_with_quality_report`（后者在几何审计基础上追加艺术分析）；不想用就直接忽略，不影响。

---

# 0.10 → 0.11 迁移指南

0.11 是**纯新增**版本：不破坏任何 0.10 的既有 API、CLI 行为或返回契约。现有代码**无需改动**。

## 新增（全部可选项）

- **PPTX 静态质量审计**：`audit_pptx(path, config=None)` 纯解析 `.pptx`（ZIP+XML），
  不开 PowerPoint、不依赖 Microsoft Office，检查越界 / 贴边 / 重叠 / 文本溢出 / autofit 风险，
  产出 `PptxAuditReport`（text / json / markdown / html）。
- **基线回归**：`compare_pptx(baseline, candidate)` 产出 `PptxDiffReport`，
  聚合新增 / 已解决 / 变化的问题与形状增删移动缩放文本变化；`--fail-on-new` 只阻断候选
  **新增或恶化**的问题。
- **Deck 生成门禁**：`deck.render_with_report(html, output, audit_mode="report"|"strict", fail_on=...)`。
  `render()` **签名与行为完全不变**，只是新增了这个带审计的变体。
- **CLI**：`offipy audit` 子命令（参数与退出码见 [docs/audit.md](audit.md)）。

## 迁移步骤

0.10 → 0.11 **无迁移步骤**。如果 0.10 代码能跑，0.11 直接换版本号即可。

唯一注意点（非破坏）：

- `import offipy` / `import offipy.audit` **仍然不加载 python-pptx**（惰性 import 硬约束）；
  只有真正 `audit_pptx` / `compare_pptx` 解析文件时才需要，且解析依赖 python-pptx
  （`pip install offipy[deck]`）。
- `offipy audit` 的退出码语义与其它命令不同：0=未达门槛 / 1=达 `--fail-on` 或 `--fail-on-new` /
  2=参数或输入错误 / 3=依赖或解析错误。这是 **audit 子命令内部**自捕异常的结果，
  不影响其它子命令的 `OffipyError → 1` 行为。

---

# 0.9 → 0.10 迁移指南

0.10 对 PowerPoint 读 API 做了一次**破坏性重构**：旧的 `read_slide_texts()`（无参，
返回全部页的 `{index, title, body, notes}` 摘要）拆成了两个职责清晰的 API。
同时在占位符常量上修正了一个既有 Bug（P0）。

本文只讲怎么迁移；完整 API 见 [docs/api](api.md)。

## 破坏性变更：read_slide_texts 签名

### 发生了什么

| 0.9 | 0.10 |
|-----|------|
| `ppt.read_slide_texts()` → 全部页摘要 `list[dict]` | `ppt.read_slide_summary()` → 全部页摘要 `list[dict]`（语义不变） |
| — | `ppt.read_slide_texts(slide_idx, *, include_empty=False, recursive=True)` → 单页 per-shape 文本记录 `list[SlideTextRecord]` |

`read_slide_texts` 现在是**按页、按 shape** 读文本；`slide_idx` 是必填位置参数。
`read_slide_summary` 承接旧的全部页摘要行为。

### 迁移步骤

**旧用法（0.9）——要全部页摘要：**

```python
items = ppt.read_slide_texts()          # 全部页 {index,title,body,notes}
```

**改为（0.10）——同一语义：**

```python
items = ppt.read_slide_summary()        # 依旧 {index,title,body,notes}
```

**改为（0.10）——要某一页的逐 shape 文本：**

```python
records = ppt.read_slide_texts(slide_idx=1)
for r in records:
    print(r["shape_id"], r["name"], r["text"])
```

### 传错参数会发生什么

旧调用 `ppt.read_slide_texts()` 现在会抛 Python 标准错误：

```
TypeError: read_slide_texts() missing 1 required positional argument: 'slide_idx'
```

这是**刻意设计**（方案 A）：`slide_idx` 保持必填，不引入运行时拦截或告警，
让缺失调用立即显式失败，避免长期签名被污染。迁移者按上面的「改为」改写即可。

## 新增：read_slide_summary

返回逐页摘要，字段与 0.9 的 `read_slide_texts()` 输出一致：
`{"index", "title", "body", "notes"}`。

- title：标题/居中标题占位符（type 1/3）优先；否则按稳定阅读顺序回退第一个**非空**非豁免文本。
- body：正文占位符（type 2）优先；否则其余文本 shape 按阅读顺序 `"\n"` 拼接。
- **空文本 shape 一律不进 title/body**（0.10.1 行为修正：title 回退曾漏过滤空文本——header 背景矩形等带空 TextFrame 的 shape 会按阅读顺序排在真标题前抢占 title，导致所有页 title 空串，真实比赛 PPT 实证。0.10.1 起 title/body 都跳过空文本候选，对齐 `read_slide_texts` 的 `include_empty=False` 语义）。
- **豁免集**：页码/页眉/页脚/日期占位符（type 13/14/15/16）+「页码候选」
  （纯数字 AND 位于页面底部/角落 AND 尺寸较小）不进 title/body。
- 对标准标题/正文占位符页面与 0.9 行为一致；纯文本框页面为启发式摘要，
  排序稳定、语义一致，**不承诺与 0.9 逐字节一致**。

## 新增：read_slide_texts（v2 语义）

`read_slide_texts(slide_idx, *, include_empty=False, recursive=True)` 返回
第 `slide_idx` 页全部**具有文本能力**的 shape 的记录，元素类型是
[`SlideTextRecord`](#slide-text-record-数据模型)（TypedDict）。

- 只返回有 TextFrame 的 shape；图片/线条/无文本图形不在此列。
- `include_empty=True` 连文本为空的 TextFrame shape 也返回。
- `recursive=True` 递归 group 内文本；非旋转 group 子元素坐标是幻灯片绝对坐标
  （`coordinate_space="slide"`），旋转 group 内不可信（`coordinate_space="unknown"`）。
- 坐标单位恒为磅（pt）。

### SlideTextRecord 数据模型

```python
class SlideTextRecord(TypedDict):
    shape_id: int
    name: str
    text: str
    left: float
    top: float
    width: float
    height: float
    coordinate_space: Literal["slide", "unknown"]
    coordinate_unit: Literal["pt"]
    is_placeholder: bool
    placeholder_type: int | None            # PpPlaceholderType 数值
    placeholder_type_name: str | None       # 完整映射 + "unknown_{n}" 兜底
    parent_shape_id: int | None
    group_path: list[int]                   # group 祖先 shape_id 链（外层→内层）
```

> **真机行为（探针实证）**：PowerPoint COM 会**拍平嵌套 group**——打开含「外层
> grpSp 包内层 grpSp」的文件时，内层 group 不出现在对象模型中，其子元素直接成为
> 外层 `GroupItems` 成员，且 `Left/Top/Width/Height` 已换算为**幻灯片绝对坐标**。
> 因此 `group_path` 反映 COM 对象模型的实际结构，通常为**单层**（如 `[300]`）；
> `parent_shape_id` 指向直接父 group。「多层祖先链」仅在 COM 真提供嵌套 group
> 时出现（`_iter_shapes` 会正确递归），真实 PowerPoint 生成的嵌套 group 一律拍平。
> `coordinate_space="slide"` 对所有 group 子元素成立（坐标本就是幻灯片绝对坐标）。

类型可从 `from offipy import SlideTextRecord` / `PLACEHOLDER_TYPE_NAMES` 导入，
mypy 能推导字段类型（见 `tests/test_api_stub.py::test_mypy_user_example_reveals_sliderecord_types`）。

## 附带修正：占位符常量（行为变化，非 API 破坏）

0.10 修正了 `src/offipy/ppt.py` 的占位符常量（P0 Bug）：

| 常量 | 0.9（错误） | 0.10（微软官方值） |
|------|------------|--------------------|
| `PP_PLACEHOLDER_TITLE` | 13（实为 slideNumber） | 1 |
| `PP_PLACEHOLDER_CENTER_TITLE` | 14（实为 header） | 3 |
| `PP_PLACEHOLDER_SLIDE_NUMBER` | — | 13 |
| `PP_PLACEHOLDER_HEADER` | — | 14 |
| `PP_PLACEHOLDER_FOOTER` | — | 15 |
| `PP_PLACEHOLDER_DATE` | — | 16 |

**影响**：`set_title` / `set_body` / `set_notes`（`_placeholder_by_type`）在标准布局上
现在能找到真正的标题/正文占位符，行为更正确——例如对「标题+内容」布局调用
`set_title` 之前会建文本框、现在会写入标题占位符。如果你的代码依赖 0.9 的错误
常量值做占位符类型判断，请改用 `from offipy import PLACEHOLDER_TYPE_NAMES`
（完整映射，含 `unknown_{n}` 兜底）。

## 不影响

- Word/Excel 的 `read_*` 操作不变。
- `read_slide_summary` 的返回值字段与 0.9 `read_slide_texts()` 相同，存量代码
  只改方法名即可。
- 新增 op 只需改 `src/offipy/ppt.py` + `src/offipy/schema.py`，server/CLI/MCP/
  stub/文档自动派生。
