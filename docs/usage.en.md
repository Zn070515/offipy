> [中文](usage.md)

# Quick Start

## Session model

`offipy` treats an Office application as a **session**: every call reconnects to the
same Office instance through the server on port 8890. The target document is resolved
by op type:

- **Read ops** (`get_cell` / `read_range` / `read_doc_text` / `read_slide_summary` / `read_slide_texts` / `get_target` …):
  default to the **currently active** document — an explicit `doc_id` takes priority, then the
  active target set by `activate` or `new_*/open_*`, then a real-time probe of
  `ActiveWorkbook` / `ActiveDocument` / `ActivePresentation` registered into the document table
  (pure probing, never implicitly creates). An unknown or stale handle raises `TargetNotFoundError`.
- **Destructive ops** (writes / formatting / save / close, etc.): **refuse to run by default**;
  you must provide one of three — an explicit `doc_id`; or `follow_active=True` (follow the
  currently active document); or an `expected_target` binding (`{"doc_id"}` / `{"name"}` /
  `{"path"}`, combinable, resolve-once). With none present, they raise `InvalidArgumentError`
  prompting for a target — they never silently write to the currently active document.

## CLI

```bash
offipy excel new_book                      # "book1"
offipy excel new_book                      # "book2" (the new book becomes the active target)
offipy excel get_target                    # points to the latest "Workbook2"
offipy excel activate --doc_id book1       # switch the active target to book1
offipy excel set_cell --sheet 1 --cell A1 --value 100 --follow-active
offipy excel set_cell --sheet 1 --cell B1 --value 200 --doc_id book2  # explicit routing
offipy excel read_range --sheet 1 --range_addr A1:B1   # read ops default to the active target
offipy excel list_docs                     # {doc_id: {name, path, active}}
offipy excel quit
```

Destructive ops need a target: `--doc_id <doc_id>` / `--follow-active` / `--expected-target '<json>'`.
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

Read operations (`read_range` / `read_doc_text` / `read_slide_summary` / `read_slide_texts` / `list_docs`) are
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
pip install "offipy[deck]"
playwright install chromium
```

## Python API

```python
from offipy import Excel

with Excel() as x:
    book = x.new_book()                    # "book1"
    x.set_cell(1, "A1", 42, doc_id=book)   # destructive ops need an explicit doc_id
    assert x.read_range(1, "A1:A1", doc_id=book) == [[42.0]]
    x.quit()
```

The Python API returns the **raw values** of App methods (`new_book` → the `"book1"`
string, `read_range` → a 2D list). `OperationResult` is the **return contract of the
HTTP `/call`** (`{ok, operation, resource_id, message, data}`, HTTP-only); MCP returns
the `data` payload — the three entry points have different return shapes; see the
honest comparison in [api.en.md](api.en.md).

## Capability boundary: create/append vs incremental edit

offipy is strongest **from a blank slate**: creating documents, appending
paragraphs/cells/slides (`new_*` / `add_*` / `set_*`), and reading the text layer
back (`read_*`). **Incremental edits to an existing document** — moving/resizing/
deleting existing shapes, fine-tuning a textbox's position/size — are not in the
library yet: `read_slide_texts` is a read-only op; its records exist so an agent
can understand the current page's text and coordinates. To change appearance, use
the "read back → rebuild/append in a new document" flow (the HTML→PPTX pipeline
rebuilds from HTML the same way). Incremental shape editing (`read_shapes` /
`set_shape_position` / `set_shape_size` / `delete_shape`) is on the roadmap, not
yet released; when it lands, `read_shapes`'s `shape_id` will share the same data
model lineage as `read_slide_texts`.
