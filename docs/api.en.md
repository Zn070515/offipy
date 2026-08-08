> [中文](api.md)

# Return Contract and the Three Entry Points

offipy exposes the same set of operations through three entry points: the **Python API** (`Excel()/Word()/Ppt()` from `offipy.api` and the underlying App classes), **HTTP RPC** (`/call`, see `src/offipy/server.py`), and **MCP** (`offipy mcp`, see `src/offipy/mcp_server.py`). All three share the same source — the same schemas / App methods and the same domain exceptions — but they have **different return shapes**. This document describes them truthfully, without forcing them into one shape.

## What Each Entry Point Returns

| Entry point | Success return | void op | Failure |
|------|----------|---------|------|
| **Python API** | The App method's raw return value (`new_book` → a `doc_id` string, `get_cell` → the cell value, `get_target` → a dict, …) | `None` | Raises an `OffipyError` domain exception (`exceptions.py`) |
| **HTTP RPC `/call`** | An `OperationResult` dict `{ok, operation, resource_id, message, data}` (plus a `result` compatibility alias) | Same, with `data: null` | HTTP status mapped from `error_code` (400/404/409/500/502/503, see `docs/protocol.md`), body `{ok:false, error, error_code, trace, ...}` |
| **MCP tools** | The operation's `data` payload (the raw value for read ops) | The string `"ok (<op>)"` | MCP error (message sourced from the same domain exceptions) |

Example: `excel.get_cell(1, "A1")`

| Entry point | How to write it | Return |
|------|------|------|
| Python | `x.get_cell(1, "A1")` | `100` |
| HTTP | `POST /call {"app":"excel","op":"get_cell","args":{"sheet":1,"cell":"A1"}}` | `{"ok":true, "operation":"excel.get_cell", "resource_id":"excel:book:book2f9a5c5e1a2b3c4d", "message":"ok", "data":100, "result":100}` |
| MCP | `excel_get_cell(sheet=1, cell="A1")` | `100` |

## OperationResult (HTTP-Only Contract)

`OperationResult` is the **`/call` response contract of the server**, defined in `src/offipy/result.py`. It is not a unified return body for Python/MCP — the Python API returns the method's raw value directly, and MCP unpacks `data` and returns it.

Success (HTTP 200):

```json
{"ok": true, "operation": "excel.set_cell", "resource_id": "excel:book:book2f9a5c5e1a2b3c4d",
 "message": "ok", "data": null, "result": null}
```

| Field | Description |
|------|------|
| `ok` | Boolean, whether the operation succeeded |
| `operation` | Full name in `"<app>.<op>"` form, e.g. `excel.set_cell` |
| `resource_id` | `"<app>:<kind>:<doc_id>"` — identifies the document the operation acted on; **doc_id is a stable identifier within the session, not the user-editable name** (`book<hex>`/`doc<hex>`/`pres<hex>`, high-entropy and opaque, not enumerable); `null` when there is no target |
| `message` | Human-readable message (usually `"ok"` on success) |
| `data` | The operation result: the raw value for read ops, `null` for void ops |
| `result` | A compatibility alias for `data` (gradual migration for older clients) |

Failure (HTTP status code is mapped from `error_code`, see the table in `docs/protocol.md`):

```json
{"ok": false, "operation": "excel.get_cell", "resource_id": null,
 "error": "TargetNotFoundError: 没有打开的工作簿", "error_code": "target_not_found",
 "trace": ["TargetNotFoundError: 没有打开的工作簿"]}
```

`error_code` maps one-to-one to the domain exceptions (see `docs/exceptions.en.md`); clients use it to map a response back to the corresponding domain exception, keeping all three entry points in sync — **whatever the HTTP status is, the semantics are always carried by the body's `error_code`**. `ComOperationError` additionally carries the `hresult` field; a deprecated op additionally carries `warning` (see `docs/deprecation.en.md`). `trace` is a redacted exception-chain message list — no paths / line numbers / source snippets.

## Session Semantics (Target Identity)

By default an op acts on the **current active document** (ActiveWorkbook / ActiveDocument / ActivePresentation):

- `get_target`: queries the identity of the currently active target `{"app", "doc_id", "name", "path"}` (`null` if none).
- `activate(doc_id)`: sets the given document as the active target and syncs the real UI (Excel `Workbook.Activate()`, Word `Document.Activate()`, PPT activates the window containing the document); on sync failure it rolls back and raises `ComOperationError`.
- `list_docs`: truthfully returns the table of documents whose handles are registered `{doc_id: {"name", "path", "active"}}` — **no implicit enumeration** — it only reports the documents for which we hold a handle.
- A destructive op may carry `expected_target` (`{doc_id}` / `{name}` / `{path}` / combinations) for **target binding**: resolve-once — the resolved `doc_id` is injected into the method call, eliminating "validate A, execute B"; an empty object or one containing unknown keys is rejected outright (`InvalidArgumentError`), and a binding failure raises `TargetNotFoundError`.

## Security Model at a Glance

The server only listens on `127.0.0.1` and authenticates with a Bearer token (see `SECURITY.en.md`); `server stop` performs graceful shutdown via the authenticated `/shutdown`. When a token mismatch occurs or port ownership cannot be proven, it **never kills** the server — it only prompts for manual handling.
