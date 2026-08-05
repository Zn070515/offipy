> [中文](SECURITY.md)

# Security Policy

offipy drives real Microsoft Office applications through a local HTTP server. This document
describes its security model and how to handle a discovered vulnerability.

## Security model

### Resident server (port 8890)

- **scope**: the server only listens on `127.0.0.1` (loopback address) and is not exposed to the
  external network. It holds COM references to Office, so **any process that can reach the port
  can drive your current Office documents**.
- **Bearer token authentication**: a random token is generated at startup, preferring the
  environment variable `OFFIPY_SERVER_TOKEN` (never written to disk); otherwise it is read,
  generated and persisted at `user_data_dir()/token` in the user data directory. Every request
  must carry the `Authorization: Bearer <token>` header, otherwise it is rejected with 401.
  **Failure to write the token is a startup failure** — this eliminates "server half-alive,
  client guaranteed 401".
- **Whitelist**: derived from the operation schema (`schema.py`) as the single source of truth
  (`server._OPS` is `frozenset(schema.ops(app))`), exposing only each app's public RPCs;
  session-internal methods (`active_pres` / `active_doc` / `active_book`) and private methods
  are never exposed. **Adding a new RPC only touches `schema.py` in one place** — the three entry
  points (server / CLI / MCP) stay in sync automatically.
- **Abuse prevention**: request body limit of 16MB (413 over the limit), `Content-Type` must be
  `application/json` (otherwise 415), negative `Content-Length` is rejected (400), response body
  limit of 64MB (over the limit returns 500 and never writes an oversized payload to the client),
  `POST` only accepts the `/call` and `/shutdown` paths, everything else returns 404.
- **Health endpoints**: `/ping` requires no auth (handshake only), `/status` requires auth and
  returns `{version, protocol, pid, python, started_at, targets}` — `targets` is a read-only
  identity snapshot of each App's currently active document (`{app, doc_id, name, path}`), cached
  by worker threads; handlers never touch COM. `/shutdown` requires auth — identity is proven by
  the token and shutdown is graceful; it does not rely on force-killing by pid.
- **Loopback binding**: only `127.0.0.1` / `localhost` / `::1` are allowed by default; `--host
  0.0.0.0` requires an explicit `--unsafe-allow-remote`, otherwise startup is refused
  (`ServerStartError`).
- **Process management (ownership discipline)**: `offipy server status|stop|restart` manage the
  resident process via a PID file + netstat probe.
  - `status` is **read-only**: when not running it only reports "not running", **never implicitly
    starts** the server.
  - `stop`: if identity can be authenticated (token matches) → authenticated `/shutdown` graceful
    shutdown; token mismatch (`auth_fail`) → **never kills**, only prompts to fix the token; port
    occupied by a non-offipy process (`mismatch`) → only force-kills when `server.pid` proves the
    process is "our" server, otherwise refuses and prompts for manual handling.
  - An old client connecting to a new server (token mismatch) only reports an error, **does not
    kill its own process**.

### Session semantics

The server keeps Office windows and documents alive, and ops act on the user's currently active
document. **Don't leak the token to untrusted processes** — having the token is equivalent to
read/write access to your current Office session.

**Target binding (`expected_target`)**: destructive ops may carry `expected_target`, keyed by
`doc_id` / `name` / `path` (combinable). At dispatch the server does resolve-once: it resolves the
target doc_id via `get_target(doc_id=...)`, validates that `name`/`path` match, then injects the
resolved result into the method parameters (eliminating "validates A, executes B"); empty objects
or objects with unknown keys are rejected (`InvalidArgumentError`), and binding failure raises
`TargetNotFoundError`. Ops bind to a target and **do not follow the user's focus**, preventing
concurrent scripts from accidentally mutating an unintended document after the focus has been
switched away. Read-only ops do not perform binding validation.

### Data on disk

- token: `user_data_dir()/token` (`%LOCALAPPDATA%\offipy` on Windows, `~/.local/share/offipy` on
  other platforms)
- converter data: overridable via `OFFIPY_CONVERTER_DATA_DIR`, defaults to the above
- feedback learning: `~/.offipy/feedback.jsonl`

## Reporting vulnerabilities

If you find a security issue (not just an ordinary bug), do not disclose the details in a public
issue, to avoid exploitation before a fix is available. Please privately contact the repository
maintainers (see the `pyproject.toml` authors) or file a **Private vulnerability report** on
GitHub (Repository → Security → Report a vulnerability).

Please include: the affected version, reproducible steps, impact assessment (for example
"an unauthorized process can read/write Office documents"), and a fix suggestion if possible.
