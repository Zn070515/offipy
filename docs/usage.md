# 快速上手

## 会话模型

`offipy` 把 Office 应用当成一个**会话**：每次调用通过 8890 端口的 server 重连同一个
Office 实例。目标文档由以下规则解析：

1. **显式 `doc_id`**：操作参数里的 `doc_id` 直接路由到指定文档（`book1` / `doc1` / `pres1`）；
   未知或已失效句柄抛 `TargetNotFoundError`。
2. **缺省活动目标**：走 `activate` 或 `new_*/open_*` 设定的活动文档。
3. **重连兜底**：无登记活动目标时，实时探测 `ActiveWorkbook` / `ActiveDocument` /
   `ActivePresentation` 并入文档表（纯探测，绝不隐式创建）。

## CLI

```bash
offipy excel new_book                      # "book1"
offipy excel new_book                      # "book2"（新书成为活动目标）
offipy excel get_target                    # 指向最新创建的"工作簿2"
offipy excel activate --doc_id book1       # 切换活动目标到 book1
offipy excel set_cell --sheet 1 --cell A1 --value 100
offipy excel set_cell --sheet 1 --cell B1 --value 200 --doc_id book2  # 显式路由
offipy excel read_range --sheet 1 --range_addr A1:B1
offipy excel list_docs                     # {doc_id: {name, path}}
offipy excel quit
```

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
offipy deck outline deck.html             # 生成大纲
```

`render` 使用**原子替换**：先写同目录临时文件，后处理完成后 `os.replace` 覆盖目标；
任何失败不会破坏已存在的 `.pptx`。转换管线需要 `offipy[deck]` 与 chromium：

```bash
uv pip install -e ".[deck]"
uv run playwright install chromium
```

## Python API

```python
from offipy import Excel

with Excel() as x:
    book = x.new_book()                    # "book1"
    x.set_cell(1, "A1", 42)
    assert x.read_range(1, "A1:A1") == [[42.0]]
    x.quit()
```

所有返回都经 `OperationResult` 统一封装：`{ok, operation, resource_id, message, data}`，
`data` 是操作的实际返回值。
