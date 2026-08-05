> [中文](deprecation.md)

# Deprecation Policy (P2-9)

offipy deprecates outdated RPC operations (ops) declaratively: set a flag in the schema, and
the server automatically adds a `warning` to responses, so consumers can gradually migrate
instead of each maintaining their own convention.

## How to Mark

Add `deprecated=True` to the `OpSpec` of the corresponding op in `src/offipy/schema.py`:

```python
"old_op": OpSpec(
    description="…",
    destructive=True,
    deprecated=True,  # ← 已弃用
),
```

Marking something deprecated only requires a single change to the schema; the server response
behavior takes effect from it (the server whitelist, CLI, and MCP registration all derive from
the schema, so there is no need to keep three places in sync).

## Server Behavior

Both **success and failure responses of a deprecated op carry the `warning` field**
(`server._deprecation_warning`):

```json
{"ok": true, "operation": "word.old_op", "data": null,
 "warning": "word.old_op 已弃用（deprecated），将在未来版本移除"}
```

Ops that are not deprecated do not carry this field. The `warning` is additional information and
does not change the HTTP status code or `error_code`; when clients/MCP receive it, they should
prompt the user to switch to the replacement op.

## Lifecycle

1. **Mark as deprecated**: set `deprecated=True` to enter the deprecation period. The replacement
   op is already available in the schema (its description notes the replacement).
2. **Notice period**: keep the op for at least one full MINOR version so downstream consumers can
   be informed and migrate via the `warning` field. The op's functionality is unchanged during this
   period.
3. **Removal**: at the next incompatible version (per SemVer, bump MINOR during the 0.x stage),
   remove the op — the App method, schema entry, tests, and documentation are cleaned up together,
   and the server whitelist narrows accordingly. The removal is recorded in `CHANGELOG.md`.

## Consumer Contract

- **client**: besides returning `data`, `call` should surface the `warning` if the response
  contains one (e.g., the CLI prints it to stderr).
- **MCP**: tool metadata may carry a deprecation notice; when the call result contains a `warning`,
  forward it to the host.
- **CLI**: prints the `warning` when running a deprecated op, without blocking the call.

## Current Status

Currently `schema.py` has **no** op marked `deprecated` — this policy is a mechanism reserved for
the future. When the first deprecated op appears, it will be recorded in `CHANGELOG.md`.

Implementation: `server._deprecation_warning` / `_success_result` / `_error_result`
(in `src/offipy/server.py`); consistency is covered by
`tests/test_server_security.py::test_deprecated_op_gets_warning`.
