> [English](assets.en.md)

# 资源系统（offipy.assets）

offipy v0.14 引入统一的**资源系统（Asset System v1）**：把图标、纹理、原生图元等
「视觉素材」抽象成 `asset://` 资源，经过 **provider 解析 → 确定性测量 → 占位符
注入 → 渲染落位** 一条管线，最终变成 PowerPoint 里**可编辑的原生对象**（freeform
矢量 / SVG picture / 原生形状与文本框），而不是贴死的位图。

- **确定性**：同一份 HTML + 同一主题，每次渲染结构一致；资源声明有稳定 ID。
- **可编辑**：图标是 native freeform，图元是原生形状 + 文本框（双击可改文字）。
- **可溯源**：每次渲染产出 `assets.json` 用法清单，记录每个资源的 provider /
  许可证 / 来源，满足合规审计。
- **无网络**：图标 / 纹理 / 图元全部 vendored 进 wheel，渲染过程不联网。

## 快速开始

```html
<div data-asset="asset://ph/icon/check"
     style="position:absolute;left:80px;top:140px;width:120px;height:120px;color:var(--accent)"></div>
<div data-asset="asset://procedural/pattern/topography?seed=42" data-asset-placement="background"
     style="position:absolute;left:0;top:0;width:100%;height:100%"></div>
<div data-asset="asset://primitives/primitive/metric-badge"
     data-asset-param-value="24%" data-asset-param-label="YoY" data-asset-param-delta="+3.2%"
     style="position:absolute;left:100px;top:200px;width:420px;height:220px"></div>
```

```bash
offipy deck make --html deck.html --theme mckinsey --layouts --out deck.pptx
```

## v0.14 内置 provider

| provider | kind | 内容 | 许可证 | first_party |
|----------|------|------|--------|-------------|
| `ph` | `icon` | 1512 个 Phosphor 图标（vendored） | MIT | 否 |
| `lu` | `icon` | 1756 个 Lucide 图标（vendored） | ISC | 否 |
| `procedural` | `pattern` | 8 个确定性生成纹理（wave / blob / dot-grid / square-grid / rings / topography / circuit / gradient-orb） | MIT | 是 |
| `primitives` | `primitive` | 8 个可编辑原生图元（quote-mark / section-number / label-pill / metric-badge / timeline-node / process-arrow / device-frame / browser-mockup） | MIT | 是 |

`AssetKind` 预留 `illustration` / `map` / `flag` 等 kind 供未来外部 provider
（v0.15 里程碑），v0.14 不实现。

## URI 语法

```text
asset://<provider>/<kind>/<name>[?k=v&k2=v2...]
```

- `<provider>` / `<kind>` / `<name>` 均小写、短横线分隔。
- 查询参数 `k=v`：键做 `_`→`-`、小写、排序的**规范化**，键必须唯一；值做
  percent-decode。十六进制颜色以 `%23RRGGBB` 往返（`#` → `%23`）。
- 拒绝 `#` 片段与 CSS `var(` 引用（参数值不得引用 CSS 变量）。
- 示例：

```text
asset://ph/icon/check
asset://lu/icon/settings
asset://procedural/pattern/topography?seed=42
asset://procedural/pattern/rings?count=4
asset://primitives/primitive/metric-badge
```

### 渲染模式

| 模式 | 含义 |
|------|------|
| `freeform_svg` | 源 SVG 解析为 PowerPoint native freeform（`p:sp` + `a:custGeom`），双击显示可编辑路径（图标） |
| `svg` | SVG 以 OOXML SVG picture 写入：主 `a:blip` 挂 PNG 栅格回退，`asvg:svgBlip` 指矢量（procedural 纹理；无 Playwright 时降级纯 SVG） |
| `svg_template` | 模板 SVG 经颜色插槽（如 `__ACCENT__`）实体化后写为 SVG picture |
| `raster` | 位图 payload（`add_picture`） |
| `native_shape` | 原生形状 + 文本框 + freeform 组合（图元） |

### 落位（placement）

| 值 | 语义 |
|----|------|
| `replace`（默认） | 渲染产物插入原声明所在 DOM 槽位，替换占位符 |
| `decorative` | 同 replace 槽位语义，标记为装饰元素 |
| `background` | 移到 `grpSpPr` 之后、所有内容形状之前（沉底） |

> **v0.14 限制**：`background` 只对 procedural 纹理等矢量资源可用；原生图元
> （primitive）声明 `data-asset-placement="background"` 会**显式报错**。

## HTML 用法

### 声明资源

```html
<div data-asset="asset://<uri>"
     data-asset-param-<key>="<value>"
     data-asset-placement="replace|decorative|background"
     style="..."></div>
```

- `data-asset-param-*`：按 URI 里 provider 的 schema 传参；十六进制颜色用
  `%23` 转义（如 `data-asset-param-accent="%232251ff"`），或直接给主题语义 token
  （`accent` / `surface` / `ink` / `muted` / `bg`）。
- 资源尺寸取自元素的测量矩形（浏览器渲染的真实 BCR），渲染产物会缩放进该矩形。

### 旧 `data-icon` 兼容

v0.12+ 的图标容器写法保持不变，内部迁移到 `ph` / `lu` provider：

```html
<svg data-icon="ph:check-circle" viewBox="0 0 256 256" width="72" height="72"></svg>
```

`data-icon="<集>:<名字>"` 等价于 `asset://ph/icon/<名字>`（或 `lu`），v0.14 无行为
变化，无需迁移。

### `data-primitive` 语法糖

原生图元提供简写：

```html
<div data-primitive="metric-badge" data-asset-param-value="24%"></div>
```

预处理后会展开成规范的 `data-asset="asset://primitives/primitive/metric-badge"` +
内部 ID，与规范写法完全等价。

### 声明注入的前置检查

资源注入依赖 visual audit 产出的 `measurements.json`。因此：

- `no_visual_audit=True`（CLI `--no-visual-audit`）与 `data-asset` /
  `data-primitive` / `data-icon` 声明**不兼容**，会在启动 chromium / convert 前
  fail-fast。
- visual-audit 渲染结束后，输出目录（如 `out_audit/`）内出现 `assets.json`；
  `no_visual_audit` 不产出资产清单。

## Python API

```python
from offipy.assets import get_default_registry

r = get_default_registry()

# 搜索
metas = r.search("grid", kind="pattern")      # 按 kind 过滤
print([m.ref for m in metas])

# 解析
asset = r.resolve("asset://procedural/pattern/topography?seed=42")
print(asset.meta.ref, asset.provider_meta.license)

# 带必填参数的图元：用 AssetRequest 传参（示例用只含可选参数的图元）
from offipy.assets import AssetRef, AssetRequest

resolved = r.resolve(AssetRequest(AssetRef("primitives", "primitive", "browser-mockup")))
assert resolved.payload.primitive == "browser-mockup"
```

`import offipy.assets` 是纯标准库表面（model / uri / registry / license / color /
materialize），**不加载** python-pptx / Pillow / playwright。

## assets.json 用法清单

visual-audit 渲染后，输出目录的 `_audit/assets.json` 记录本次渲染用到的每个资源：

```json
{
  "schema": 1,
  "assets": [
    {
      "declaration_id": "asset-s01-001",
      "slide_index": 1,
      "request": "asset://primitives/primitive/metric-badge",
      "provider": {"id": "primitives", "license": "MIT", "source_url": "...", "first_party": true},
      "placement": "replace"
    }
  ]
}
```

- `declaration_id`：稳定、确定性递增（`asset-s<slide>-<seq>`）。
- 图表（`data-chart`）**不是** Asset System v1 资源，不会出现在清单里。

## 原生图元 schema（v0.14）

所有图元共有的可选参数：`accent`（默认主题 accent）、`fill`（图元专属默认）。
参数在 provider 层校验，非法参数显式报错。

| 图元 | 参数 | 默认填充 | 结构 |
|------|------|----------|------|
| `quote-mark` | `text`（必填，≤240） | transparent | 引号字形 + 文本框 |
| `section-number` | `number`（必填 int 0..9999）、`label`（可选 ≤120） | surface | 数字 + 可选标签 + accent 装饰 |
| `label-pill` | `text`（必填，≤120） | accent | 圆角矩形 + 居中文案 |
| `metric-badge` | `value`（必填 ≤80）、`label`（可选 ≤120）、`delta`（可选 ≤40） | surface | 卡片 + value/label/delta |
| `timeline-node` | `label`（可选 ≤120）、`phase`（可选 past/current/future，默认 current） | — | 节点圆点 + 标签；phase 决定样式 |
| `process-arrow` | `steps`（必填，逗号分隔 2..8 项，每项 ≤80）、`direction`（可选 horizontal/vertical，默认 horizontal） | surface | 箭头分 N 段，每段文本可编辑 |
| `device-frame` | `device`（必填 phone/tablet/desktop） | surface | 原生设备外框 + 空屏 |
| `browser-mockup` | `title`（可选 ≤120）、`url`（可选 ≤240） | surface | 窗口卡片 + 顶部 chrome + 三圆点 |

已知限制（v0.14 明示，不视为 bug）：

- `process-arrow` 的 `steps` 不支持转义逗号；标签含逗号请换别的图元。
- `device-frame` / `browser-mockup` **不接受截图/嵌套资源**：没有 `screenshot` /
  `src` / `image` 参数，屏幕是空的可编辑区域。
- 图元文字用默认安全 sans 字体（Arial 系）；主题字体 token 是 v0.15 方向，v0.14
  不引入新的公开字体 token 模型。

## 非目标 / 未来 provider

- v0.15 里程碑：外部插画 / 地图 / 照片类 provider（`illustration` / `map` /
  `flag`），需要网络或大体积素材，v0.14 不做。
- 嵌套资源（如「设备帧里放截图」）是显式**非目标**，遇到即报错而不是静默降级。
