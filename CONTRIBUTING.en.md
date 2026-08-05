> [中文](CONTRIBUTING.md)

# Contributing

offipy is a Windows-only Office COM automation library. Fixes and improvements are welcome, but
please follow the discipline below first — it exists so that every commit is reviewable,
revertible, and never breaks ongoing session-based automation.

## Development environment

- Windows + Git Bash. Commands use bash syntax.
- Use `uv` to manage `.venv` (Python 3.12):
  `uv run python ...` / `uv run pytest ...` / `uv run offipy ...`
- `export PYTHONIOENCODING=utf-8` before invoking Python subprocesses, otherwise Chinese text
  garbles in the Windows terminal.

## Branch discipline (highest priority)

- All feature/fix/docs development starts by creating a branch first; never commit development
  code directly to main.
- Naming: `feat/<kebab-case>` (feature), `fix/<kebab-case>` (fix), `docs/<kebab-case>` (docs),
  `build/<kebab-case>` (build/CI/dependencies).
- Confirm main is clean before branching: `git status` shows no uncommitted changes →
  `git checkout main` → `git checkout -b <branch-name>`.
- The only exception that may commit directly to main: **emergency fix** (CI down, broken build,
  blocking bug); state why you didn't use a branch in the commit.
- Merging: after all checks pass on the branch, `git checkout main && git merge <branch> --no-ff`
  (or via PR), then `git branch -d <branch>`.

## Commit conventions

- Small incremental commits: specify concrete files, a clear message, prefix `feat:` / `fix:` /
  `docs:` / `build:` / `chore:` + a one-line description.
- Before committing, run `git status` / `git diff` to confirm only this task's changes are
  included, with no unrelated files.
- Single source of truth for the version: `__version__` in `src/offipy/__init__.py`; version
  bumps are their own commit (`chore: bump version to X.Y.Z`), never mixed with feature changes.
- Before the initial stable release (1.0.0), every version's first component is 0; breaking
  changes bump MINOR, not MAJOR.
- **Pre-release numbering strategy**: before the initial stable release, TestPyPI uses pre-release
  numbers such as `0.9.0a1` / `0.9.0rc1`; for a stable release, `__version__`, the git tag, and the
  CHANGELOG top level must all match (backed by an alignment test). **Do not repeatedly bump an
  unpublished version number** — the 0.9.0 stable release was never published, so pre-release
  fixes still iterate as `0.9.0a1` / `0.9.0rc1` without bumping the version further.

## Gates (all must be green before merging)

```bash
uv sync --extra dev
uv run ruff check .
uv run ruff format --check .
uv run mypy src/offipy
uv run pytest tests -q
```

`src/offipy/_vendor/` is vendored third-party code (the HTML→PPTX converter) and is not subject to
lint/format/mypy constraints.

## Adding an RPC: only touch schema.py

The `OPS` table in `src/offipy/schema.py` is the **single registration source** for ops — the
server whitelist (`server._OPS`), CLI argument type coercion (`cli._coerce_kwargs`), and MCP tool
registration (`mcp_server._build_tool`) are all derived from it. Adding an op takes three steps:

1. Implement the method in the corresponding App class (`excel.py` / `word.py` / `ppt.py`);
2. Register the `OpSpec` in `schema.py` (`params` / `returns` / `readonly` / `destructive` /
   `description`);
3. The three entry points are ready automatically — **do not** hand-write `server._OPS` or MCP
   decorators.

Consistency is enforced by `tests/test_schema_consistency.py` (bidirectional consistency between
schema ↔ `server._OPS`, every schema op has a corresponding MCP tool, readonly/destructive flags
match).

## Exception policy A (domain exceptions)

Library-level failures always raise `OffipyError` subclasses, **never `SystemExit` / bare COM
exceptions**:

- `InvalidArgumentError` (`invalid_argument`): invalid argument/input (parsing, constant tables,
  range validation)
- `TargetNotFoundError` (`target_not_found`): no open workbook/document/presentation, or
  `expected_target` binding failed
- `FileConflictError` (`file_conflict`): the target file already exists and `overwrite` was not
  explicitly given
- `ComOperationError` (`com_operation`): a COM call failed inside an App method (keeps `hresult`
  for disconnection detection)
- `ProtocolError` (`protocol`): protocol mismatch

Each exception carries a `code` that maps one-to-one to the RPC failure response `error_code`. The
CLI catches at the boundary and converts it to an exit code + stderr message; the MCP server
catches and converts it to a tool error response; library callers can catch directly.

## Required after verifying Office integration

After each test/verification passes, proactively close any Office windows that were opened and
confirm the processes have exited:

```bash
uv run offipy quit ppt     # /excel /word; quit is just obj.Quit() and may leave residuals
tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"   # if residuals remain → taskkill //F //PID <pid>
```

After modifying a module the server depends on (`ppt.py` / `server.py` / `mcp_server.py`, etc.),
you must restart the server process on 8890, otherwise the server loads stale code.

## Notes on testing the server

- `tests/test_server_security.py` starts a real HTTPServer (temporary port) to verify
  authentication / request limits / the op whitelist — it does not dispatch real ops, so **no
  Office is needed**.
- A 401 does not kill the server: a failed token check only rejects the request; the server keeps
  serving `/ping` and calls with a correct token.
- `offipy server status|stop|restart` manage the resident process (PID file + netstat probe); after
  changing server code, restart with `offipy server restart`, and when verifying, confirm there is
  no Office residual in `tasklist`.
- The in-memory tests in `tests/test_mcp_server.py` do not spawn subprocesses: they use the mcp SDK
  in-memory streams (`create_client_server_memory_streams`) + `_lowlevel_server.run` to connect
  directly to the protocol layer, verifying tool registration / structured returns / annotations /
  error mapping — they run without Office and are the preferred regression for MCP changes.
- Non-Windows collection: the COM test modules' `pytestmark` uses a
  `sys.platform != "win32" or not core.running(...)` short-circuit plus the `com` marker, so
  collection does not blow up on Linux and tests skip automatically. Deck conversion tests carry
  the `deck_render` marker.

## Asset-driven development (asset libraries / icon libraries / materials)

First do `WebSearch` research + use `gh` to check existing GitHub solutions (open-source libraries,
licenses, maintenance status, maturity), don't build from scratch behind closed doors. After
choosing, record an ADR with the sources, respect licenses, and note the provenance in the asset
files.
