> [English](usage.en.md)

# 快速上手

## 会话模型

`offipy` 把 Office 应用当成一个**会话**：每次调用通过 8890 端口的 server 重连同一个
Office 实例。目标文档按 op 类型解析：

- **读 op**（`get_cell` / `read_range` / `read_doc_text` / `read_slide_summary` / `read_slide_texts` / `get_target` …）：
  缺省作用在**当前激活**文档上——优先显式 `doc_id`，其次 `activate` 或 `new_*/open_*` 设定的
  活动目标，再次实时探测 `ActiveWorkbook` / `ActiveDocument` / `ActivePresentation` 并入文档表
  （纯探测，绝不隐式创建）。未知或已失效句柄抛 `TargetNotFoundError`。
- **破坏性 op**（写入 / 格式 / 保存 / 关闭等）：**默认拒绝执行**，必须三选一——
  显式 `doc_id`；或 `follow_active=True`（跟随当前激活文档）；或 `expected_target` 绑定
  （`{"doc_id"}` / `{"name"}` / `{"path"}` 可组合，resolve-once）。三者都没有时抛
  `InvalidArgumentError` 提示补目标——绝不静默改到当前激活的文档。

## CLI

```bash
offipy excel new_book                      # "book<hex>"（高熵 doc_id，随机生成）
offipy excel new_book                      # 再次新建，返回另一个"book<hex>"（新书成为活动目标）
offipy excel get_target                    # 指向最新创建的"工作簿2"
offipy excel activate --doc_id book<hex>   # 切换活动目标到指定 book<hex>
offipy excel set_cell --sheet 1 --cell A1 --value 100 --follow-active
offipy excel set_cell --sheet 1 --cell B1 --value 200 --doc_id book<hex>  # 显式路由
offipy excel read_range --sheet 1 --range_addr A1:B1   # 读 op 缺省走活动目标
offipy excel list_docs                     # {doc_id: {name, path, active}}
offipy excel quit

# PPTX 静态质量审计与基线回归（纯解析，不需要打开 PowerPoint，无 Office 依赖）
offipy audit deck.pptx                     # 文本报告（默认）
offipy audit deck.pptx --fail-on HIGH      # 达 HIGH → 退出码 1（CI 门禁）
offipy audit deck.pptx --format html --out audit.html --slides-dir export/   # SVG 画布报告
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID      # 基线回归：只阻断新增/恶化
```

破坏性 op 需要一个目标：`--doc_id <doc_id>` / `--follow-active` / `--expected-target '<json>'`。
布尔参数用 `--key true/false`：`--overwrite true`。结构化值可用 `--payload '{"...": ...}'`。
参数名以下划线分隔（如 `--range_addr`、`--doc_id`），类型由 `schema.py` 声明并自动转换。
PPTX 审计的参数与退出码、Python API 详见 [docs/audit.md](audit.md)（审计）与
[docs/audit-baseline.md](audit-baseline.md)（基线回归）。

## CLI 退出码契约

| 退出码 | 语义 |
| --- | --- |
| `0` | 成功 |
| `2` | 使用 / 参数 / 预运行无效输入（`InvalidArgumentError`）——目标缺失、路径不存在、非法参数值 |
| `1` | 运行时领域失败（`OffipyError` 系：`ComOperationError` / `FileConflictError` / `TargetNotFoundError` / `RemoteCallError` …） |
| `offipy audit` | 专属契约：0=未达门槛 / 1=达 `--fail-on` / `--fail-on-new` / 2=参数或输入错 / 3=依赖或解析错 |
| `offipy deck audit` | 专属契约：0=通过 / 1=未通过 / 2=参数或输入错 |

通用命令错误统一以 `[app::op] 失败: <可读消息>` 输出到 stderr，不泄露 traceback。
`InvalidArgumentError` 同时是 `OffipyError` 与 `ValueError` 子类，CLI 捕获顺序先判
`InvalidArgumentError → 2`、再判 `OffipyError → 1`——预运行时错误不会误报为运行时失败。

## LLM 设计 → 可编辑 PPTX（Agent 原生）

offipy 把 diagram-design skill（vendored，MIT）的 LLM 设计能力以 **Agent 原生**模式
接入：offipy 自身**不调用 LLM、不 spawn agent**——把 skill 注册给宿主 agent
（Claude Code / Codex），agent 按产物契约把设计落地为 Mermaid/drawio 源码文件，
offipy 只做「产物 → 可编辑 PPTX」。

### 安装 skill

```bash
offipy diagram install_skill                  # 装到 ~/.claude/skills/（幂等，不覆盖用户编辑）
offipy diagram install_skill --target_dir <技能目录> --force   # 指定目录 / 覆盖重建
```

安装后宿主 agent 即可发现 `diagram-design`（设计指引）与 `offipy-diagram`（产物契约桥接）
两个 skill。

### 转换

```bash
offipy diagram build --source design.mmd --out design.pptx
offipy diagram build --source design.drawio --out design.pptx
```

- `source` 必须是已存在的 `.mmd` / `.drawio` 文件（不接受内联文本）
- `direction`（Mermaid 流向，LR/TB…）与 `page`（draw.io 页名/序号）按格式各自透传
- `out` 已存在时默认拒绝覆盖（`FileConflictError`，CLI 退出码 1），重新生成加 `--overwrite true`
- 输出是 16:9 整页**可编辑形状** PPTX；布局由 Mermaid/drawio 引擎重排

### 可转换子集（契约边界）

Mermaid **只支持** `flowchart/graph`、`sequenceDiagram`、`stateDiagram-v2`、`erDiagram`；
`gantt` / `journey` / `mindmap` / `timeline` / `gitgraph` 等不支持——请改用 draw.io
表达或重构为上述 kind。

Python API 等价：`offipy.diagrams.mermaid_to_pptx(source, out, direction=...)` /
`offipy.drawio.drawio_to_pptx(source, out, page=...)`；MCP：`diagram_build` /
`diagram_install_skill`。

## Server 生命周期

```bash
offipy server status     # 只读探测，未运行返回"server 未在运行"，不拉起
offipy server stop       # 鉴权 /shutdown 优雅停机
offipy server restart    # stop 后重新拉起
offipy server --port 8891   # 多实例：按端口隔离 token/pid/oplog
```

`server status` 报告协议版本、`session_id`、每应用目标身份。token 生命周期与不杀策略见
[协议](protocol.md) 与 SECURITY.md。

## MCP

MCP server 走 stdio，工具集合从 `schema.py` 自动注册。Claude Desktop 配置示例：

```json
{
  "mcpServers": {
    "offipy": {
      "command": "offipy",
      "args": ["mcp"]
    }
  }
}
```

读操作（`read_range` / `read_doc_text` / `read_slide_summary` / `read_slide_texts` / `list_docs`）标记只读，
写操作标记会改动状态；`save` / `save_pdf` 暴露 `overwrite` 参数。

## HTML→PPTX 管线（deck）

```bash
offipy deck make --html deck.html --out deck.pptx --no-open
offipy deck outline --input outline.md --out deck.html   # markdown 大纲 → HTML 骨架
```

`render` 使用**原子替换**：先写同目录临时文件，后处理完成后 `os.replace` 覆盖目标；
任何失败不会破坏已存在的 `.pptx`。首次转换会在源 HTML 旁生成 `.audited.html` 工作副本，
后续 audit 修复改副本、原 HTML 不动；**0.10.2 起源 HTML 比副本新时自动重建副本**（改源即刻
生效，不再静默复用旧副本）。HTML 中本地 `<img src>` 引用文件缺失时转换**直接报错**并列出
缺失文件（0.10.2 起，不再静默嵌入空白占位图）。转换管线需要 `offipy[deck]` 与 chromium：

```bash
pip install "offipy[deck]"
playwright install chromium
```

`open_live` 打开实况演示时先复制到系统临时目录的 `offipy-live-*` 副本再让 PowerPoint
演示——PowerPoint 锁定的是副本，产物 `.pptx` 路径不被锁，同路径 `render(overwrite=True)`
可反复重渲染（#22）。用 `deck.close_live(doc_id)` 关闭实况演示并清理副本；目标文件被
占用时 `render` 会给出可操作错误（提示先 `close_live` / `offipy quit ppt`，或换输出名）。

`data-asset` / `data-primitive` / `data-asset-param-*` / `data-asset-placement` 资源
声明与 `asset://` URI、provider 与 `assets.json` 溯源，见 [资源系统](assets.md)。

### Mermaid 图

页面里写 `<pre class="mermaid">` 块，渲染后由 `offipy.diagrams` 替换为 PowerPoint 原生
可编辑形状（复用 charts 注入管线，需 visual audit 的 measurements.json）：

```html
<section data-pptx-slide>
  <h1>部署架构</h1>
  <pre class="mermaid">graph TD
    A[构建] --> B[测试]
    B --> C[发布]
    C --> D[生产]</pre>
</section>
```

`deck render deck.html` 后 mermaid 块变为可编辑形状；或独立调用
`offipy.diagrams.mermaid_to_pptx("graph TD\nA-->B", "out.pptx")` 直接出 16:9 整页可编辑
PPTX（TD/TB/LR/RL/BT 方向、subgraph 容器、中文 label）。

注意：`graph`/`flowchart` 必须带显式方向（如 `graph TD`，裸 `graph` 会被拒）；暂不支持
`%%` 注释（vendored 解析器契约）。

### draw.io 图

页面里写 `<div class="drawio" data-drawio="arch.drawio"></div>`，路径基于该 HTML 所在目录；
`deck render` 后由 `offipy.drawio` 把 `.drawio` 转成 PowerPoint 原生可编辑形状（复用 charts
注入管线，需 visual audit 的 measurements.json），保留 draw.io 作者摆好的版式与配色：

```html
<section data-pptx-slide>
  <h1>部署架构</h1>
  <div class="drawio" data-drawio="arch.drawio"></div>
</section>
```

`deck render deck.html` 后 draw.io 块变为可编辑形状；或独立调用
`offipy.drawio.drawio_to_pptx("arch.drawio", "out.pptx")` 直接出 16:9 整页可编辑 PPTX。
`page=` 接受 int 页码（0 起）或 str 页名，默认第一页（不暴露 all）。

- 相对 `data-drawio` 路径在 deck 管线内可用：HTML 被拷到临时目录时自动改写为
  `file://` 绝对 URI，源文件仍能解析（也可直接写 `file://` URI）。
- 多页 `.drawio` 在 deck 注入里**必须**用 `data-drawio-page="N"`（1 基）指定页，
  否则报错而不是静默取第一页：
  `<div class="drawio" data-drawio="arch.drawio" data-drawio-page="2"></div>`；
  页名也可用（不区分大小写）。页号最小为 1，`all` 不支持。
- 节点 `fontSize` 按容器缩放换算成 pt（缺省 12pt），字号层级随框缩放，不被拍平。
- 正交/曲线边按折点（waypoints）渲染成折线，保留箭头；`strokeWidth`、`rotation`、
  `dashPattern` 透传到 PPTX 形状（`dashPattern` 为空格分隔的数对）。

已知限制：图标类节点（`icon:*`）兜底为矩形、draw.io 自定义形状兜底为圆角矩形。

## Python API

```python
from offipy import Excel

with Excel() as x:
    book = x.new_book()                    # "book<hex>"（高熵 doc_id）
    x.set_cell(1, "A1", 42, doc_id=book)   # 破坏性 op 需显式 doc_id
    assert x.read_range(1, "A1:A1", doc_id=book) == [[42.0]]
    x.quit()
```

Python API 返回 App 方法的**原始值**（`new_book`→`"book<hex>"` 字符串、`read_range`→二维列表）。
`OperationResult` 是 **HTTP `/call` 的返回契约**（`{ok, operation, resource_id, message, data}`，
HTTP-only），MCP 返回 `data` 载荷——三入口返回形状不同，如实对照见 [api.md](api.md)。

`Excel()/Word()/Ppt()` 本地直连 COM 对象绑定创建它的线程（STA），**非线程安全**——
同一实例须在创建线程内使用；跨线程各自建 facade 时用 `offipy.direct.com_apartment()`
包一层（线程各自 CoInitialize/Uninitialize）。连到既有实例默认不改其可见性；
确需改传 `modify_existing_visibility=True`。

PPTX 质量审计不需要 Office、不需要打开 PowerPoint：

```python
from offipy import audit_pptx, compare_pptx, Severity

report = audit_pptx("deck.pptx")
if report.max_severity is not None and report.max_severity >= Severity.HIGH:
    print("存在 HIGH 级问题，拒绝发布")
print(report.to_markdown())

diff = compare_pptx("baseline.pptx", "candidate.pptx")
if diff.gate_severity() is not None and diff.gate_severity() >= Severity.MID:
    print("候选相对基线新增/恶化 MID+ 问题")
```

### 艺术/审计证据层

- **图片拉伸失真按解码 vs 渲染判定**（#126）：`distorted_image` 的 `natural_ratio` /
  `physical_ratio` 现基于图片**解码尺寸**（img 取 `naturalWidth/Height`，SVG 按 viewBox 定比）
  与**渲染尺寸**（CSS 布局宽高）之比，能检出真实拉伸漂移；旧实现两个比值都取自渲染尺寸，
  漂移恒≈0，规则形同虚设。该语义变化使历史 FEATURES 样本作废，
  `feature_schema_version()` 由 2→3。
- **PPTX-only 富集**（#128）：只传 `pptx=` 时，`audit_pptx` / `analyze_scene` 现在从
  `_ShapeRecord` 解析字号 / 字体 / 前景色 / 背景色 / 透明度 / fill_kind，hierarchy /
  typography / color 维度的证据覆盖率从 0 提升（不再完全没有字号 / 颜色证据）；
  schemeClr / sysClr（主题色引用）仍不解析，纯像素类证据（PNG / measurements）仍需额外源。
- **元素 opacity**（#137）：measurement 的元素级透明度透传到 `ArtElement.opacity`（0-1，
  None=无证据）；PPTX-only 路径按 run 前景色 alpha 与形状填充 alpha 的 min 合并（最透明部分
  决定可见性）。
- **fill_kind 标记**（#140）：measurement 记录 `fill_kind`（`solid` / `gradient` / `shadow` /
  `image`），gradient / shadow 等光栅化装饰可被 audit 识别，不再误判为普通色块。

## feedback 学习系统（v0.18）

三层 feedback 语义（文档钉死，避免混淆）：
- `offipy.feedback`（v1）：维度权重，`dimension_weights()`，`~/.offipy/feedback.jsonl`
- `offipy.art.feedback`（v2）：规则 ±1，`recommend_adjustments` → `feedback_severity_adjustments`
- `offipy feedback`（v3，本系统）：可学习 numpy MLP，`feedback_train` / `feedback_status` / `feedback_append`

训练：`offipy feedback train`（读 `~/.offipy/art_feedback.jsonl` → 训练 →
写 `~/.offipy/art_feedback_model.json`）。样本不足/无有效样本时返回状态 JSON，
不删除已有模型（F2-E）。**数值门禁（#112）**：loss 非有限 → `training_diverged`、
输出恒定（output_std < 1e-6）→ `model_collapsed`，坏模型一律不写、原子保留旧模型，
训练做全局梯度裁剪。需要 numpy：`pip install "offipy[feedback]"`。

**学习质量（#115-#122）**：
- **预处理标准化（#118/#120）**：训练集拟合零方差特征 drop + 高相关（|r|≥0.99）去重 +
  全局 z-score，mean/scale/kept 持久化到 model.json `preprocessing` 块（model schema v2），
  推理端同一 transform
- **ensemble K=5（#122）**：多 seed member 取平均降方差 + worth 校准（quality_score 归一）；
  **abstain**（|worth| 近零 / member 分歧大的 finding 不 shift）+ **OOD**（特征 z 越界不 shift），
  保守回退 v2
- **样本级 repeated stratified CV（#119）**：按 rule 分层的重复 5 折、绝不做 pair-level split；
  95% 置信下界触 chance → `poor_generalization` soft flag（只记录不拒绝写盘）
- **容量自适应告警（#121）**：容量按独立样本数自适应（H≈√n），`samples_per_param`
  分级 ok/warn/critical
- **逐规则样本诊断（#117）**：样本不足（insufficient_pairs）时返回 per-rule
  `{fixed, accepted, pairs, single_direction, suggest}` 可行动建议
- tiny_image 特征补全 + FEATURES schema v1→v2 bump（#115）

追加标签：`offipy feedback append --profile <p> --rule_id <r> --action fixed
--severity MID --features '<json>' --feedback_dir <dir>`（写入该目录 JSONL，
供 train 学习；severity 限 LOW/MID/HIGH）。

状态：`offipy feedback status`（样本数 / 有效样本 / 配对潜力 / 模型 none|valid|expired；
模型有效时另表面 `effective_dims` / `samples_per_param` / `poor_generalization`）。

消费侧要求（#113）：学习**消费**必须显式 `feedback_dir`——`analyze_scene(feedback=True)`
不带目录 → `InvalidArgumentError`；`learned_adjustments` 不带目录 → 返回 None，
不静默加载全局 `~/.offipy` 模型。写入侧（train/append）仍默认 `~/.offipy`。

CLI 学习消费通道（#114/#116）：
- `offipy deck audit --feedback-dir <dir>`：审计时应用反馈学习（加载指定目录记录/模型）；
  **带 `--feedback-dir` 时报告 / `--json` 输出透传 `experimental_score`**（#116）
- `offipy deck make --export-png <dir>`：导出 PNG 反馈目录（旧名 `--feedback` 为弃用别名，
  语义不变，仍只是导出目录，不触发学习消费）

推理消费点：
- `rule.delta.<rule_id>`：历史记录 worth 均值 → ±1 调整 → `feedback_severity_adjustments`
- `finding.severity_shift`：analyze 后处理 pass，仅 rule-computed（无 override）finding 生效；
  **按规则证据门禁（#111）**——仅该规则有效标签 ≥3 才被 shift，0 标签规则不做跨规则泛化；
  模型不确定（abstain）/ 特征 OOD 的 finding 不 shift（保守回退 v2）
- `quality.score`：替换 `experimental_score`（仅 `include_experimental_score=True` 时算）；
  同证据门禁，只由通过门禁（可被 shift）的 finding 贡献；值由 ensemble 均值 worth 经校准
  （worth_scale 归一）映射

冷启动：无模型 / 模型过期（feature_schema_version / model schema 不符）/ 未装 numpy →
完全回退 v2 行为（`recommend_adjustments`）。删除 model.json 即回到 v2。学习系统是可拆卸增强，
绝不回退 audit 硬门禁。

## 能力边界：创建/追加 vs 增量修改

offipy 擅长**从无到有**：新建文档、追加段落/单元格/页（`new_*` / `add_*` / `set_*`），
以及把文本层读回（`read_*`）。对**既有文档的增量修改**——移动/缩放/删除已存在的
shape、精修某个文本框的位置尺寸——目前不在库内：`read_slide_texts` 是只读操作，
返回结构只用于 Agent 了解当前页的文本与坐标；要改外观时请走「读回 → 在新文档上
重建/追加」的流程（HTML→PPTX 管线从 HTML 重建也是同一思路）。增量 shape 编辑
（`read_shapes` / `set_shape_position` / `set_shape_size` / `delete_shape`）已列入
路线图，尚未发布；届时 `read_shapes` 的 shape_id 与 `read_slide_texts` 的数据模型
一脉相承。
