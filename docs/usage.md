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
offipy excel new_book                      # "book1"
offipy excel new_book                      # "book2"（新书成为活动目标）
offipy excel get_target                    # 指向最新创建的"工作簿2"
offipy excel activate --doc_id book1       # 切换活动目标到 book1
offipy excel set_cell --sheet 1 --cell A1 --value 100 --follow-active
offipy excel set_cell --sheet 1 --cell B1 --value 200 --doc_id book2  # 显式路由
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

## Python API

```python
from offipy import Excel

with Excel() as x:
    book = x.new_book()                    # "book1"
    x.set_cell(1, "A1", 42, doc_id=book)   # 破坏性 op 需显式 doc_id
    assert x.read_range(1, "A1:A1", doc_id=book) == [[42.0]]
    x.quit()
```

Python API 返回 App 方法的**原始值**（`new_book`→`"book1"` 字符串、`read_range`→二维列表）。
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

## 能力边界：创建/追加 vs 增量修改

offipy 擅长**从无到有**：新建文档、追加段落/单元格/页（`new_*` / `add_*` / `set_*`），
以及把文本层读回（`read_*`）。对**既有文档的增量修改**——移动/缩放/删除已存在的
shape、精修某个文本框的位置尺寸——目前不在库内：`read_slide_texts` 是只读操作，
返回结构只用于 Agent 了解当前页的文本与坐标；要改外观时请走「读回 → 在新文档上
重建/追加」的流程（HTML→PPTX 管线从 HTML 重建也是同一思路）。增量 shape 编辑
（`read_shapes` / `set_shape_position` / `set_shape_size` / `delete_shape`）已列入
路线图，尚未发布；届时 `read_shapes` 的 shape_id 与 `read_slide_texts` 的数据模型
一脉相承。
