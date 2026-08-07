> [中文](exceptions.md)

# Exception Contract

Failures at the library layer always raise a catchable `OffipyError` subclass
(strategy A: full domain exceptions), never `SystemExit`. The three entry points
(Python API / RPC / MCP) share the same origin: every exception carries a `code`, the
server's failure response carries an `error_code`, and the client maps it back to the
corresponding domain exception per the table below.

## Exception table

| Exception | `code` | Triggered when |
| --- | --- | --- |
| `OffipyError` | `offipy` | Base class of all exceptions |
| `InvalidArgumentError` | `invalid_argument` | Arguments/input are invalid: cell parsing, constant tables, range validation fail |
| `TargetNotFoundError` | `target_not_found` | Target does not exist: no workbook/document/presentation open; unknown `doc_id`; `expected_target` binding mismatch |
| `FileConflictError` | `file_conflict` | Target file already exists and `overwrite=True` was not explicit (also inherits `FileExistsError`) |
| `ComOperationError` | `com_operation` | A COM call inside an App method fails; `hresult` preserves the underlying HRESULT for disconnection detection |
| `ProtocolError` | `protocol` | Request/response protocol version mismatch, or the handshake fails |
| `OfficeUnavailableError` | `office_unavailable` | Office application/COM runtime is unavailable |
| `ServerStartError` | `server_start` | The local resident server cannot start or times out on startup |
| `RemoteCallError` | `remote_call` | A remote call to the resident server fails (op error/timeout/network error) |
| `ConversionError` | `conversion` | HTML→PPTX conversion/rendering fails (including a missing chromium) |
| `UnsupportedPlatformError` | `unsupported_platform` | A Windows-only capability is called on a non-Windows platform |

## Compatibility inheritance

`InvalidArgumentError` also inherits `ValueError`, and `FileConflictError` also inherits
`FileExistsError`: existing callers using `except ValueError` / `except FileExistsError`
need no changes.

## Boundary handling

- **CLI**: converts exceptions to exit codes at the boundary — `InvalidArgumentError` → 2
  (usage / argument / pre-runtime invalid input), other `OffipyError` subclasses → 1 (runtime
  domain failure); stderr is cleaned before output (no leaked tracebacks). `offipy audit`
  keeps its dedicated 0/1/2/3 contract and `deck audit` keeps 0/1 (see
  [docs/usage.en.md](usage.en.md) "CLI exit-code contract").
- **MCP server**: catches them and returns them to the model as tool errors
  (`is_error: true`).
- **Library callers**: simply catch the corresponding domain exception directly.

## expected_target

Destructive operations (`destructive=True` in the schema) support `expected_target`
target binding; the key is taken from `doc_id` / `name` / `path` (combinable). The
server resolves **resolve-once** at dispatch time: it resolves the target `doc_id` via
`get_target(doc_id=...)`, validates that `name`/`path` match, and then injects the
resolved result directly into the method call parameters (eliminating "validate A,
execute B"). An empty object or one containing unknown keys → `InvalidArgumentError`; a
binding failure (target does not exist / name/path mismatch) →
`TargetNotFoundError`. The binding target is the **bound target**, which does not follow
user focus, preventing destructive operations from running against the wrong document.
