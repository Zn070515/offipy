> [English](audit.en.md)

# PPTX 质量审计

`offipy audit` 是一套**静态几何质量门禁**：不打开 PowerPoint、不依赖 Microsoft Office，
直接解析 `.pptx`（ZIP+XML）的结构化 shape 提取，检查越界 / 贴边 / 重叠 / 文本溢出 /
autofit 风险，产出 text / json / markdown / html 报告，并按严重度门槛阻断不合格产物。

- 输出**稳定的 `rule_id`**（不是自然语言 message）——用户 / CI 依赖它做自动化。
- 可配置豁免、可读 suppressed（为什么没报）、可读 warnings（什么解析不了）。
- 与 `compare_pptx`（[基线回归](audit-baseline.md)）配合做「新改动有没有引入新问题」的回归。
- 与 `deck render_with_report` 配合做「HTML→PPTX 生成即门禁」。

## 安装与依赖

审计核心**不依赖 Microsoft Office**（纯解析，无 COM）。读取 `.pptx` 需要 python-pptx：

```bash
pip install "offipy[deck]"    # 含 python-pptx
```

`import offipy` / `import offipy.audit` **不加载** python-pptx（惰性 import 硬约束）；
只有真正解析文件时才需要。

## 快速开始

```bash
# 文本报告（默认）
offipy audit deck.pptx

# 报告 + 门槛：达 HIGH 就退出码 1（CI 用）
offipy audit deck.pptx --fail-on HIGH

# JSON / Markdown / 单文件 HTML（SVG 画布，可筛选）
offipy audit deck.pptx --format markdown
offipy audit deck.pptx --format html --out audit.html --slides-dir export/
```

```python
from offipy import audit_pptx

report = audit_pptx("deck.pptx")
print(report.max_severity)      # Severity.HIGH / MID / LOW / None
print(report.to_markdown())
for f in report.findings:
    print(f.rule_id, f.severity.name, f.message)
```

## CLI 参考

```
offipy audit <file.pptx>
  --format text|json|markdown|html    默认 text（html 缺省写 <stem>.audit.html）
  --out PATH                          报告输出文件；text/json/markdown 缺省打 stdout
  --fail-on HIGH|MID|LOW             审计达该严重度 → exit 1（默认 HIGH）
  --baseline PATH                    给出则走回归对比（见 audit-baseline.md）
  --fail-on-new HIGH|MID|LOW         对比模式：候选新增/恶化达该严重度 → exit 1
  --safe-margin FLOAT                安全边距（英寸，默认 0.2）
  --bounds-tolerance FLOAT           越界容差（英寸，默认 0.01）
  --no-full-bleed-ignore             关闭全页背景豁免
  --no-repeated-decoration-ignore    关闭重复装饰豁免
  --no-page-number-ignore            关闭页码豁免
  --no-header-footer-ignore          关闭页眉页脚豁免
  --slides-dir PATH                  html 专用：PNG 页面背景目录（slide-<n>.png）
  --show-suppressed                  豁免项默认已列出；本标志保留兼容
  --debug                            失败时打印完整 traceback
```

**退出码**：

| 码 | 含义 |
|----|------|
| 0 | 未达门槛（审计通过） |
| 1 | 成功但达 `--fail-on` / `--fail-on-new`（门槛命中） |
| 2 | 参数或输入错误（文件不存在、`--fail-on-new` 误用等） |
| 3 | 依赖或解析错误（缺 python-pptx、ZIP/XML 损坏） |

`_audit_main` 自捕全部预期异常转成退出码，绝不与其它命令的 `OffipyError → 1` 冲突。

## Python API

```python
from offipy import (
    Severity, AuditConfig, AuditFinding,
    PptxAuditReport, PptxDiffReport, audit_pptx, compare_pptx,
)

report = audit_pptx(
    "deck.pptx",
    AuditConfig(
        safe_margin_in=0.2,          # 安全边距（英寸）
        bounds_tolerance_in=0.01,     # 越界容差（英寸）
        ignored_shapes={(1, 42)},     # 用户显式豁免：第 1 页 shape #42
        ignored_regions=[(0.0, 6.5, 10.0, 1.0)],  # 底部区域整体豁免（英寸 x,y,w,h）
    ),
)
```

### Severity

`LOW=1 / MID=2 / HIGH=3`（`IntEnum`）。**比较必须按整数值**（`f.severity >= Severity.HIGH`），
禁止字符串比较。序列化输出为 `"LOW"/"MID"/"HIGH"`。

### 稳定 rule_id

| rule_id | 严重度 | 含义 |
|---------|--------|------|
| `geometry.bounds.partial` | MID / HIGH | 形状部分越出幻灯片边界 |
| `geometry.bounds.off_canvas` | LOW / MID | 完全在画布外（暂存/动画/设计残留） |
| `geometry.margin.left/right/top/bottom` | LOW | 内容贴近边缘，间距 < 安全边距 |
| `geometry.overlap.partial` | LOW / MID | 形状部分重叠（覆盖比 > 0.5，以较小形状计） |
| `geometry.overlap.covered_text` | MID / HIGH | 一个形状完全覆盖另一个；文本被图片/图表盖住为 HIGH |
| `text.fit.horizontal` | LOW / MID | 文本横向超出文本框（显式 `wrap="none"` 单行超宽） |
| `text.fit.vertical` | LOW / MID | 文本纵向超出文本框（显式多行超高 / 无可用空间） |
| `text.autofit.shrink` | MID / HIGH | normAutofit 缩小字体，可能低于最小可读 8pt |
| `text.autofit.grow` | MID / HIGH | spAutoFit 扩大 Shape，可能越界/撞对象 |

### 报告模型

- `PptxAuditReport`：`max_severity`（无 finding 为 `None`）、`findings` / `suppressed` /
  `warnings` / `shapes`（逐 shape 几何快照）、`to_dict()` / `to_json()` / `to_markdown()` /
  `to_html(slides_dir=...)`。JSON 输出**完全安全**（无 Enum / Path / set）。
- `AuditFinding`：`rule_id` / `kind` / `severity` / `message` / `primary` / `secondary`
  （两个 shape 引用）/ `details` / `confidence`。message 是中文自然语言，rule_id 才是稳定键。
- 坐标单位：**幻灯片绝对英寸**（group 子元素已绝对化）。

## 规则说明

规则注册表驱动（`DEFAULT_RULES`），依次执行：Bounds → Margin → Overlap → TextFit → Autofit。

### Bounds（越界）

- 任一边超出页面且超 `bounds_tolerance_in` → `geometry.bounds.partial`。
- 超出比例大（`out_ratio > 0.5` 或最大超出 > 0.25×页面长边）→ **HIGH**，否则 **MID**。
- 与画布无交集 → `geometry.bounds.off_canvas`（面积 ≥1 in² 为 MID，否则 LOW）。
- 已报越界的边**不再报同方向 margin**（防双报）。
- 隐藏 / group / 几何无法解析的对象跳过。

### Margin（贴边）

- `safe_margin_in=0.2`；普通内容间距 < 0.2 → `geometry.margin.*` LOW。
- **豁免**（进 suppressed 带 reason）：全页背景（`full_bleed`）、页码（`page_number`）、
  页眉页脚（`header_footer`）、重复装饰（`repeated_decoration`）、用户 ignore
  （`user_shape` / `user_region`）。
- 连接线 / 隐藏 / group 跳过。

### Overlap（重叠）

- 每页 bbox 快拒 + O(n²)；`ratio = 重叠面积 / min(两形状面积)`，> 0.5 才报。
- 跳过：连接线 / 隐藏 / group / 全页背景 / 极小装饰点（面积 < 0.0025 in²）/
  父子·祖孙对 / 几何无法解析。
- **Pair 分类**：文本位于填充 AutoShape 内 → 卡片容器，抑制为
  `intentional_containment`；文本被图片/图表完全覆盖 → HIGH；同 Group 且 z-order
  合理 → 降严重度。
- `covered_text` 只在**被盖形状有文本**时发射——空框 / 装饰点浮空卡片（双方无文本）
  不报。
- 旋转形状用 AABB 近似（confidence 0.5，message 标注「旋转包围盒近似」）。

### TextFit（文本溢出）

- 可用区域**先减 TextFrame 的 margin**。
- 横溢**仅对显式 `wrap="none"` 的文本框**报（单行不折行）；`square`（含 bodyPr@wrap
  未设，PowerPoint 默认自动折行）永不报横溢。段落含 `a:br` 软换行时取**最长段**宽，
  不跨段求和。超高按显式行数×行高；行高读取段落 `a:lnSpc`（`spcPts` 绝对 / `spcPct`
  百分比），未设时回退 `字号×1.2`；行尾软换行（`a:br`）不计多余空行。
- 超宽 / 超高需**超过 1pt 噪声下限**才报（Pillow FreeType 与 PowerPoint DirectWrite
  度量引擎存在亚 pt 级差异）。
- 字体度量：**优先 fontTools 解析字体文件**（`.ttf` / `.ttc`，如微软雅黑
  `msyh.ttc`/`msyhbd.ttc`），按 hmtx 字宽 + kerning/GPOS 字距求和 → confidence 0.8；
  失败再回退 **Pillow** 的 `getlength`，最后**字符权重**（CJK=1.0 / ASCII=0.5 /
  space=0.35 → confidence 0.4，message 标注「字符估算低置信」）。
- 页码 / 页眉 / 页脚小文本跳过（本就紧凑）。

### Autofit（自适应风险）

两种模式**分开**（不统一降级）：
- `normAutofit`（缩小字体适应 Shape）→ `text.autofit.shrink`：记录原始字号 /
  fontScale / 估算后字号；估算后 < 8pt → HIGH。
- `spAutoFit`（扩大 Shape 适应文字）→ `text.autofit.grow`：撑大后可能越界 → HIGH，
  否则 MID。

## 角色与豁免（suppressed）

`suppressed` 是**有 reason 的豁免记录**，不静默丢弃。常见 reason：

| reason | 触发 |
|--------|------|
| `full_bleed` | 全页背景（覆盖 ≥90% + 近中心 + 低 z-order + 无文本） |
| `page_number` | 纯数字 + 底部 15% 区域 + 小尺寸（或 slide number 占位符） |
| `header_footer` | 页眉页脚占位符；或跨页重复 >60% + 顶部/底部区域 |
| `repeated_decoration` | 指纹重复 >60% 页的装饰 |
| `intentional_containment` | 文本位于填充 AutoShape 内（卡片容器） |
| `decorative_overlay` | 实心小装饰/色条浮有文本容器（尺寸判别）——不遮挡内容 |
| `text_on_background` | 文字浮无文本背景/容器（非包含）——下方无内容可遮挡 |
| `transparent_overlay` | 透明（`a:noFill`）无文本上层——视觉上不遮挡任何内容 |
| `decorative_layering` | 双方无文本的长条装饰分层（短边≥3 倍、较长一维 ≥70% 大者、非包含）——部分重叠但不遮挡内容 |
| `user_shape` / `user_region` | 用户通过 `ignored_shapes` / `ignored_regions` 显式豁免 |

overlap 遮挡判定按「上层是否真的盖住下层内容」：透明无文本上层不遮挡、
小装饰尺寸判别、文字浮无文本背景不遮挡；**有文本的上层（透明与否）一律不豁免**——
文本叠文本是真问题，透明不改变内容冲突。

**不会全局忽略所有纯数字短文本**——只有底部小尺寸的纯数字才算页码。

## confidence 语义

| confidence | 含义 |
|------------|------|
| 1.0 | 精确几何，无启发式 |
| 0.8 | fontTools 字体度量（hmtx 字宽 + kerning/GPOS 字距，含 `.ttc` 集合） |
| 0.5 | 旋转形状的 AABB 近似（message 标注） |
| 0.4 | 字符权重回退（message 标注「字符估算低置信」） |

## warnings（解析异常）

`warnings` 记录解析层无法精确处理的项：`group.no_transform`（group 缺 `a:xfrm` →
子元素无法精确定位 → 跳过需精确位置的规则）、`audit.extract.slidesize_corrupt`
（演示文稿级 `sldSz@cx/@cy` 非数字 → 幻灯片尺寸降级 0.0，仍继续提取）。遇到这些
情况**不会**悄悄按零旋转处理。

## 已知限制与误报控制

- **text-fit 不执行以下结构的检查**（静默跳过、不报死、不误报）：表格单元格 /
  SmartArt / 图表内部文本 / WordArt / 竖排 / 复杂项目符号与自定义行距。结构化
  unsupported warning（`textfit.table_unsupported` 等）暂未实现，不承诺时间线。
- 旋转组 / flip 的 bounds/margin/overlap 用 AABB 近似（几何形状本身占位不变，
  翻转只影响内容朝向）。
- 不承诺所有 PPT 零误报——固定验收集（`tests/fixtures/audit/`）保证
  connector / hidden / rotate / flip / group 不误判、全页背景 / 页码不误报 margin、
  合理卡片包含不误报 overlap；其余靠 `ignored_shapes` / `ignored_regions` / `--no-*-ignore`
  微调。

## Deck 生成门禁（render_with_report）

HTML→PPTX 渲染后即审计，按模式决定放行：

```python
from offipy import deck

result = deck.render_with_report(
    "deck.html", audit_mode="strict", fail_on=deck.Severity.HIGH,
)
# 通过：替换目标并返回 RenderResult（output_path + audit_report）
# 未过：抛 deck.AuditGateError（report 在异常上、临时文件已清理、旧目标不动）
```

```bash
offipy deck make --html deck.html --out deck.pptx --no-open \
  --audit-mode strict --fail-on HIGH --audit-report deck.audit.json
```

- `report`（默认）：生成 → 审计 → 替换 → 返回 `RenderResult`。
- `strict`：生成 → 审计 → 最高严重度 ≥ `fail_on` → 抛 `AuditGateError`
  （报告先落盘、旧 `.pptx` 不被破坏）；未达 → 替换。
- 原子替换：转换先写同目录临时文件，审计通过才 `os.replace`。

## CI 用法

```bash
# 阻断：任何 HIGH 问题
offipy audit deck.pptx --fail-on HIGH

# 回归：只阻断候选新增/恶化的 MID+ 问题（基线已有的历史问题放行）
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID
```

用 `--format json` 把结果交给下游；`rule_id` 是稳定机器键，`message` 只给人类看。
