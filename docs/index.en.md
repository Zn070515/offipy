> [中文](index.md)

# offipy

**offipy** is a Windows-only Office COM automation library: session-based
driving of Word / Excel / PowerPoint, plus an HTML→editable PPTX conversion pipeline.
It treats an Office application as a **session**, reuses the same instance across
processes, and provides three consistent entry points: server / CLI / MCP.

## Features

- **Session-based**: reuses the same Office instance across processes; `ActiveWorkbook` /
  `ActiveDocument` / `ActivePresentation` resolve the target in real time.
- **Multi-document (P2-2)**: each application maintains a document table, `doc_id`
  routes explicitly, and `activate` / `list_docs` manage the active target; multiple
  server instances are isolated by port.
- **The three entry points are consistent**: HTTP server (127.0.0.1:8890, Bearer token),
  CLI (`offipy <app> <op>`), and MCP tools, all derived from a single source of truth in
  `schema.py`.
- **Domain exceptions**: `InvalidArgumentError` / `TargetNotFoundError` / `FileConflictError` /
  `ComOperationError` / `ProtocolError` all inherit from `OffipyError`; RPC carries `error_code`.
- **HTML→PPTX pipeline**: Chromium renders the HTML layout → editable PPTX, with support
  for charts, icons, and themes.
- **diagram-design skill integration (Agent-native)**: `offipy diagram build` turns
  Mermaid / draw.io sources that a host agent designed against the artifact contract into
  editable PPTX; `offipy diagram install_skill` installs the design guide plus the
  contract-bridge skill into the host agent's skill directory.

## Installation

```bash
py -m pip install "offipy[all]"     # office + deck + mcp
py -m playwright install chromium   # required by the deck pipeline
```

Split by use case: `offipy[office]` (COM automation only), `offipy[deck]` (HTML→PPTX),
`offipy[mcp]` (MCP server).

## Quick start

```bash
offipy excel new_book            # returns "book<hex>" (high-entropy doc_id)
offipy excel set_cell --sheet 1 --cell A1 --value 42 --follow-active
offipy excel read_range --sheet 1 --range_addr A1:A1   # [[42.0]]
offipy excel quit
```

To connect MCP to Claude Desktop, see the MCP configuration section of the README,
pointing at `offipy mcp`.

## Documentation

- [Quick start](usage.en.md): complete usage of sessions, multi-document, CLI / server / MCP.
- [API reference](api/index.en.md): all operations (auto-generated from schema).
- [Exception contract](exceptions.en.md): strategy A domain exceptions and the RPC `error_code` mapping.
- [Protocol](protocol.en.md): HTTP protocol, token, `/shutdown`, protocol version handshake.
- [Compatibility matrix](compatibility.en.md): Windows / Office / Python / extras support.
- [Deprecation policy](deprecation.en.md): deprecation workflow and the response `warning` field.
- [Migration guide](migration.md): 0.9 → 0.10 breaking changes and migration steps.
- [Asset System](assets.en.md): `asset://` icons / textures / native primitives, providers and `assets.json` provenance.

## Building the docs

```bash
uv run --with mkdocs-material mkdocs build
```

`docs/api/` is auto-generated from `schema.py` by `scripts/gen_api_ref.py`; rerun it
after modifying operations.
