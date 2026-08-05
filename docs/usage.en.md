> [中文](usage.md)

# Quick Start

## Session model

`offipy` treats an Office application as a **session**: every call reconnects to the
same Office instance through the server on port 8890. The target document is resolved
by the following rules:

1. **Explicit `doc_id`**: a `doc_id` in the operation parameters routes directly to the
   given document (`book1` / `doc1` / `pres1`); an unknown or stale handle raises
   `TargetNotFoundError`.
2. **Default to the active target**: falls back to the active document set by `activate`
   or `new_*/open_*`.
3. **Reconnect fallback**: when no active target is registered, probes
   `ActiveWorkbook` / `ActiveDocument` / `ActivePresentation` in real time and registers
   it in the document table (pure probing, never implicitly creates).

## CLI

```bash
offipy excel new_book                      # "book1"
offipy excel new_book                      # "book2" (the new book becomes the active target)
offipy excel get_target                    # points to the latest "Workbook2"
offipy excel activate --doc_id book1       # switch the active target to book1
offipy excel set_cell --sheet 1 --cell A1 --value 100
offipy excel set_cell --sheet 1 --cell B1 --value 200 --doc_id book2  # explicit routing
offipy excel read_range --sheet 1 --range_addr A1:B1
offipy excel list_docs                     # {doc_id: {name, path, active}}
offipy excel quit
```

Boolean parameters use `--key true/false`: `--overwrite true`. Structured values can be
passed with `--payload '{"...": ...}'`. Parameter names use underscores (e.g.
`--range_addr`, `--doc_id`); types are declared in `schema.py` and converted automatically.

## Server lifecycle

```bash
offipy server status     # read-only probe; if not running returns "server not running" without starting it
offipy server stop       # authenticated /shutdown, graceful shutdown
offipy server restart    # starts it again after stop
offipy server --port 8891   # multiple instances: token/pid/oplog isolated by port
```

`server status` reports the protocol version, `session_id`, and the target identity of
each application. For the token lifecycle and the no-kill policy, see
[Protocol](protocol.en.md) and SECURITY.md.

## MCP

The MCP server uses stdio, and the tool set is registered automatically from
`schema.py`. Example Claude Desktop configuration:

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

Read operations (`read_range` / `read_doc_text` / `read_slide_texts` / `list_docs`) are
marked read-only; write operations are marked as modifying state; `save` / `save_pdf`
expose the `overwrite` parameter.

## HTML→PPTX pipeline (deck)

```bash
offipy deck make --html deck.html --out deck.pptx --no-open
offipy deck outline --input outline.md --out deck.html   # markdown outline → HTML skeleton
```

`render` uses **atomic replacement**: it first writes a temporary file in the same
directory, then `os.replace`s it over the target after post-processing completes; any
failure will not corrupt an existing `.pptx`. The conversion pipeline requires
`offipy[deck]` and chromium:

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

The Python API returns the **raw values** of App methods (`new_book` → the `"book1"`
string, `read_range` → a 2D list). `OperationResult` is the **return contract of the
HTTP `/call`** (`{ok, operation, resource_id, message, data}`, HTTP-only); MCP returns
the `data` payload — the three entry points have different return shapes; see the
honest comparison in [api.en.md](api.en.md).
