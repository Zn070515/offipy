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

- title：标题/居中标题占位符（type 1/3）优先；否则按稳定阅读顺序回退第一个非豁免文本。
- body：正文占位符（type 2）优先；否则其余文本 shape 按阅读顺序 `"\n"` 拼接。
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
