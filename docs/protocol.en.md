> [中文](protocol.md)

# offipy HTTP Protocol (P2-8)

offipy drives real Office through a local HTTP server (default `127.0.0.1:8890`). This document defines the request/response contract, authentication, and version handshake of the `offipy-http/v1` protocol. The implementation lives in `src/offipy/server.py` (server side) and `src/offipy/client.py` (client side).

## Version Handshake

- Protocol name constant: `offipy-http/v1` (`server._PROTOCOL` / `client.PROTOCOL`).
- Request side: `/call` and `/shutdown` must carry the request header `X-Offipy-Protocol: offipy-http/v1`.
  - Missing or mismatched → 400 with `error_code: "protocol"` (mapped back to `ProtocolError`).
  - This is a **request-side handshake**: when an old client connects to a new server (or vice versa), protocol negotiation fails rather than silently mismatching.
- Server side: the `result.protocol` of the `/status` response reports the server's protocol version, letting the client probe it (`client._probe` uses this to decide whether the server is "ours").

## Endpoints

| Endpoint | Method | Auth | Description |
|------|------|------|------|
| `/ping` | GET | No | Health check; returns `{"ok": true, "result": "pong"}` and exposes no data |
| `/status` | GET | Yes | Read-only snapshot of process / protocol / session identifier / target identity |
| `/call` | POST | Yes | Executes one RPC operation (COM ops are enqueued on the worker queue and run serially) |
| `/shutdown` | POST | Yes | Authenticated graceful shutdown (after replying, shutdown is triggered by a separate thread; no reliance on force-killing by pid) |

All other paths return 404.

## Authentication

All protected endpoints require `Authorization: Bearer <token>`. Token sources: the `OFFIPY_SERVER_TOKEN` environment variable takes precedence; otherwise a persistent `user_data_dir()/token` file is used. A validation failure yields only 401 and does **not kill the server** (a token mismatch is a configuration problem, not a process problem).

## /call Request

```json
POST /call
Authorization: Bearer <token>
Content-Type: application/json
X-Offipy-Protocol: offipy-http/v1

{"app": "excel", "op": "set_cell", "args": {"sheet": 1, "cell": "A1", "value": 100}}
```

| Field | Description |
|------|------|
| `app` | `excel` / `word` / `ppt` (unknown → 400 `invalid_argument`) |
| `op` | Must be an op in the schema allowlist (`server._OPS` is derived from `schema.py`; unknown → 400 `invalid_argument`) |
| `args` | Keyword arguments passed through to the App method; path-like arguments (`path`/`out`, etc.) are absolutized by the client against the caller's CWD |

Boundary checks (all fail fast in the handler thread, never touching COM):
- `Content-Type` must be `application/json`, otherwise 415.
- The request body is capped at 16MB; exceeding it returns 413.
- A negative `Content-Length` → 400 (`read(negative)` would swallow the connection buffer).

## /call Response

Responses follow the OperationResult contract (`src/offipy/result.py`). **This is an HTTP-only contract** — the Python API and MCP each have their own return shapes (Python returns the method's raw value; MCP returns the `data` payload). For a comparison across the three entry points, see [`docs/api.en.md`](api.en.md).

Success (HTTP 200):

```json
{"ok": true, "operation": "excel.set_cell", "resource_id": "excel:book:book1",
 "message": "ok", "data": null, "result": null}
```

- `resource_id`: `"<app>:<kind>:<doc_id>"` identifies the document the operation acted on (the raw COM object is never exposed; `resource_id` is used instead); **doc_id is a stable identifier within the session, not the user-editable name**; `null` when there is no target.
- `data`: the operation result (the raw value for read ops; `null` for void ops).
- `result`: a compatibility alias for `data` (gradual migration for older clients).

Failure (HTTP 500):

```json
{"ok": false, "operation": "excel.set_cell", "resource_id": null,
 "error": "TargetNotFoundError: 没有打开的工作簿", "error_code": "target_not_found",
 "trace": ["..."]}
```

- `error_code` maps one-to-one to the domain exceptions (see `exceptions.py`):

| error_code | Domain exception |
|------------|----------|
| `invalid_argument` | `InvalidArgumentError` |
| `target_not_found` | `TargetNotFoundError` |
| `file_conflict` | `FileConflictError` |
| `com_operation` | `ComOperationError` (preserves the `hresult` field) |
| `protocol` | `ProtocolError` |
| `internal` | No mapping (unknown / ordinary exceptions) |

- The client maps a response back to the corresponding domain exception via `error_code`, keeping the three entry points (Python/RPC/MCP) in sync.

## Additional Fields

- **warning**: when an op is marked `deprecated` in the schema, both success and failure responses carry a `warning` field (see `docs/deprecation.en.md`).
- **destructive confirmation**: destructive ops carrying the `overwrite` parameter get unified overwrite protection via `paths.ensure_writable` — if the target file already exists and `overwrite=false`, → `FileConflictError` (`error_code: "file_conflict"`). `expected_target` provides target binding for destructive ops: **resolve-once** — the three keys `{doc_id}` / `{name}` / `{path}` (combinable) resolve the target via `get_target(doc_id=...)`; after validating name/path, the resolved `doc_id` is injected into the method call (eliminating "validate A, execute B"); an empty object or one containing unknown keys is rejected outright as `invalid_argument`, and a binding failure is `target_not_found` (see `SECURITY.en.md`).

## /status Response

```json
{"ok": true, "result": {
  "version": "0.9.0a1",
  "protocol": "offipy-http/v1",
  "session_id": "<uuid4>",
  "pid": 28776,
  "python": "3.12.10",
  "started_at": 1785858307.49,
  "targets": {"excel": null, "word": null, "ppt": {"app": "ppt", "doc_id": "pres1", "name": "...", "path": "..."}}
}}
```

- `session_id`: the current server session identifier (uuid4), written into every operation log entry (`oplog.jsonl`), for distinguishing / tracing across instances.
- `targets`: a read-only identity snapshot of each App's currently active document, cached by the worker thread; the handler never touches COM, so `GET /status` never launches Office as a side effect of probing.

## Operation Log (P2-3)

After every `/call`, the server appends a JSONL entry to `user_data_dir()/oplog.jsonl`:

```json
{"ts": "...", "session_id": "<uuid4>", "app": "excel", "op": "set_cell",
 "ok": true, "error_code": null, "duration_ms": 12, "resource_id": "excel:book:book1"}
```

`args` are never written to disk (sanitized/redacted). The log rotates at ~5MB (keeping a `.1` backup). Reading: `offipy log` / `offipy log --tail N`.
