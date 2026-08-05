> [English](audit-baseline.en.md)

# PPTX 基线回归（compare_pptx）

配合 [审计](audit.md) 做「**新改动有没有引入新问题**」的回归：拿一份**基线** PPTX 与一份
**候选** PPTX 对比，聚合新增 / 已解决 / 变化的问题，以及形状增删 / 移动 / 缩放 / 文本变化。

- **基线**：上次审核通过的产物（或历史版本）。
- **候选**：本次改动后的产物。
- 候选**相对基线**「新增或恶化」的问题才触发门禁；基线里已有的历史问题**放行**。

## CLI

```bash
# 回归对比：新增/恶化达到 MID 就退出码 1
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID

# 回归 + HTML 报告
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID \
  --format html --out regression.html
```

- `--baseline PATH`：给出后进入对比模式（缺省 `--format text`）。
- `--fail-on-new HIGH|MID|LOW`：**只针对候选新增或恶化的问题**。达该严重度 → 退出码 1。
- `--baseline` 必须搭配 `--fail-on-new`（只给 `--baseline` 缺 `--fail-on-new` → 退出码 2）；
  `--fail-on-new` 没有 `--baseline` 也是误用 → 退出码 2。
- 对比模式退出码与普通审计一致：0=未达门槛 / 1=门槛命中 / 2=参数或输入错 / 3=依赖或解析错。

## Python API

```python
from offipy import compare_pptx, AuditConfig

diff = compare_pptx(
    "baseline.pptx",
    "candidate.pptx",
    audit_config=AuditConfig(safe_margin_in=0.2),
)

print(diff.gate_severity())   # Severity.MID / LOW / HIGH / None（无新增/恶化）
for f in diff.added_findings:
    print("新增", f.rule_id, f.severity.name)
for f in diff.resolved_findings:
    print("已解决", f.rule_id)
print(diff.to_json())         # 完全 JSON 安全
```

## Shape 匹配链

每页按序匹配，**前一级命中即匹配**（`compare.py` 的 `_match_slide`）：

1. 同页相同 `shape_id`；
2. `name + shape_type`；
3. 归一化文本 hash + 中心距离（≤1.0 英寸取最近）；
4. 图片内容 sha256（仅 PICTURE 可取得）。

仍未匹配 → 计入 `unmatched_baseline` / `unmatched_candidate`（低置信，由调用方决定是否告警）。

> **Shape ID 不能假设跨所有编辑永久稳定**——它只是第一优先匹配键，插入/删除会让后续
> shape_id 错位，因此必须靠后面的链逐级兜底。

## PptxDiffReport 字段

| 字段 | 含义 |
|------|------|
| `baseline_path` / `candidate_path` | 双路径 |
| `baseline_sha256` / `candidate_sha256` | 双文件指纹 |
| `baseline_slide_count` / `candidate_slide_count` | 页数（`added_slides` / `removed_slides` 为差值属性） |
| `baseline_findings` / `candidate_findings` | 双方各自的完整审计结果 |
| `added_findings` | 候选新增的 Finding |
| `resolved_findings` | 基线有、候选已消失的 Finding |
| `changed_findings` | 形状已匹配、同一 Finding 严重度变化（`ChangedFinding.worsened` 标记上升） |
| `added_shapes` / `removed_shapes` | 形状增删 |
| `moved_shapes` / `resized_shapes` / `text_changes` | 位置 / 尺寸 / 文本变化（带旧新几何，英寸） |
| `unmatched_baseline` / `unmatched_candidate` | 匹配链兜不住的对象（低置信） |
| `warnings` | 双方解析异常并集 |

**门禁语义**：
- `new_or_worsened` → 候选新增或恶化的 Finding 列表（`--fail-on-new` 只看这些）。
- `gate_severity()` → 候选新增或恶化的最高严重度；无则 `None`（门槛不触发）。

```python
# 手工判断：候选是否引入 MID+ 的新增/恶化
if diff.gate_severity() is not None and diff.gate_severity() >= Severity.MID:
    raise SystemExit("候选引入 MID+ 新增/恶化问题")
```

## 判定细节

- 严重度比较**按整数值**（`Severity` 是 `IntEnum`），禁止字符串比较。
- `changed_findings` 里 `worsened=True` 才算「恶化」，进入 `new_or_worsened` 与 `gate_severity`；
  严重度**下降**的 change 不触发门禁（只记录）。
- Finding 的匹配键是 `(rule_id, primary(slide,shape_id), secondary)`；主/次形状有任一未匹配
  则视为新增。

## 固定验收集（tests/fixtures/audit/）

- `baseline.pptx` + `candidate.pptx` 由生成脚本产出：候选新增越界 `bounds.partial` MID、
  已解决贴边 `margin.right`、移动 / 缩放 / 文本变化、形状增删。
- `tests/test_audit_fixtures.py::test_compare_finds_added_and_resolved` 断言：
  `added_findings` 含 `bounds.partial` MID；`resolved_findings` 含 `margin.right`；
  `added_shapes==2` / `moved_shapes==2` / `resized_shapes==1` / `text_changes==1`；
  `gate_severity()==MID`。

## CI 用法

```bash
# 只阻断候选新增/恶化的 MID+ 问题（基线已有的历史问题放行）
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID

# 更严：任何新增/恶化即阻断
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new LOW

# 机器可读：把 diff 交给下游
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID --format json
```

> `rule_id` 是稳定机器键；`message` 只给人类看。`--format json` 的 `max_new_severity`
> 即 `gate_severity()` 的序列化名（`"LOW" / "MID" / "HIGH" / null`）。
