> [English](usage.en.md)

# 快速上手

## 会话模型

`offipy` 把 Office 应用当成一个**会话**：每次调用通过 8890 端口的 server 重连同一个
Office 实例。目标文档按 op 类型解析：

- **读 op**（`get_cell` / `read_range` / `read_doc_text` / `read_slide_texts` / `get_target` …）：
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
```

破坏性 op 需要一个目标：`--doc_id <doc_id>` / `--follow-active` / `--expected-target '<json>'`。
布尔参数用 `--key true/false`：`--overwrite true`。结构化值可用 `--payload '{"...": ...}'`。
参数名以下划线分隔（如 `--range_addr`、`--doc_id`），类型由 `schema.py` 声明并自动转换。

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

读操作（`read_range` / `read_doc_text` / `read_slide_texts` / `list_docs`）标记只读，
写操作标记会改动状态；`save` / `save_pdf` 暴露 `overwrite` 参数。

## HTML→PPTX 管线（deck）

```bash
offipy deck make --html deck.html --out deck.pptx --no-open
offipy deck outline --input outline.md --out deck.html   # markdown 大纲 → HTML 骨架
```

`render` 使用**原子替换**：先写同目录临时文件，后处理完成后 `os.replace` 覆盖目标；
任何失败不会破坏已存在的 `.pptx`。转换管线需要 `offipy[deck]` 与 chromium：

```bash
pip install "offipy[deck]"
playwright install chromium
```

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
