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
offipy excel new_book                      # "book<hex>" (high-entropy doc_id, random)
offipy excel new_book                      # creating again returns another "book<hex>" (new book becomes active)
offipy excel get_target                    # points to the latest "Workbook2"
offipy excel activate --doc_id book<hex>   # switch the active target to the given book<hex>
offipy excel set_cell --sheet 1 --cell A1 --value 100 --follow-active
offipy excel set_cell --sheet 1 --cell B1 --value 200 --doc_id book<hex>  # explicit routing
offipy excel read_range --sheet 1 --range_addr A1:B1   # read ops default to the active target
offipy excel list_docs                     # {doc_id: {name, path, active}}
offipy excel quit

# PPTX static quality audit & baseline regression (pure parsing, no PowerPoint, no Office dependency)
offipy audit deck.pptx                     # text report (default)
offipy audit deck.pptx --fail-on HIGH      # exit code 1 at HIGH (CI gate)
offipy audit deck.pptx --format html --out audit.html --slides-dir export/   # SVG canvas report
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID      # baseline regression: only block added/worsened
```

Destructive ops need a target: `--doc_id <doc_id>` / `--follow-active` / `--expected-target '<json>'`.
Boolean parameters use `--key true/false`: `--overwrite true`. Structured values can be
passed with `--payload '{"...": ...}'`. Parameter names use underscores (e.g.
`--range_addr`, `--doc_id`); types are declared in `schema.py` and converted automatically.
For the audit flags, exit codes, and Python API see [docs/audit.en.md](audit.en.md) (audit) and
[docs/audit-baseline.en.md](audit-baseline.en.md) (baseline regression).

## CLI exit-code contract

| Exit code | Meaning |
| --- | --- |
| `0` | Success |
| `2` | Usage / argument / pre-runtime invalid input (`InvalidArgumentError`) — missing target, path not found, invalid argument value |
| `1` | Runtime domain failure (`OffipyError` family: `ComOperationError` / `FileConflictError` / `TargetNotFoundError` / `RemoteCallError` …) |
| `offipy audit` | Dedicated contract: 0=threshold not reached / 1=gates at `--fail-on` / `--fail-on-new` / 2=argument or input error / 3=dependency or parse error |
| `offipy deck audit` | Dedicated contract: 0=pass / 1=fail / 2=argument or input error |

Generic command errors are emitted to stderr as `[app::op] 失败: <readable message>`, without leaking
tracebacks. `InvalidArgumentError` is a subclass of both `OffipyError` and `ValueError`, so the CLI
checks `InvalidArgumentError → 2` before `OffipyError → 1` — pre-runtime errors are never misreported
as runtime failures.

## LLM design → editable PPTX (Agent-native)

offipy wires the LLM design capability of the vendored diagram-design skill (MIT) in
**Agent-native** mode: offipy itself **never calls an LLM and never spawns an agent** —
it registers the skill with the host agent (Claude Code / Codex), and the agent writes
the design to a Mermaid / draw.io source file per the artifact contract; offipy only
does the "artifact → editable PPTX" conversion.

### Installing the skill

```bash
offipy diagram install_skill                  # installs to ~/.claude/skills/ (idempotent, never overwrites user edits)
offipy diagram install_skill --target_dir <skills-dir> --force   # custom dir / force rebuild
```

After installation the host agent can discover two skills: `diagram-design` (design
guidance) and `offipy-diagram` (artifact-contract bridge).

### Converting

```bash
offipy diagram build --source design.mmd --out design.pptx
offipy diagram build --source design.drawio --out design.pptx
```

- `source` must be an existing `.mmd` / `.drawio` file on disk (inline text is not accepted)
- `direction` (Mermaid flow direction, LR/TB…) and `page` (draw.io page name/index) are
  forwarded per format
- An existing `out` is refused by default (`FileConflictError`, CLI exit code 1); re-generate
  with `--overwrite true`
- The output is a 16:9 full-page **editable-shape** PPTX; layout is re-flowed by the
  Mermaid / draw.io engine

### Convertible subset (contract boundary)

Mermaid **only supports** `flowchart/graph`, `sequenceDiagram`, `stateDiagram-v2`,
`erDiagram`; `gantt` / `journey` / `mindmap` / `timeline` / `gitgraph` etc. are not
supported — express them with draw.io instead, or refactor into one of the supported kinds.

Python API equivalent: `offipy.diagrams.mermaid_to_pptx(source, out, direction=...)` /
`offipy.drawio.drawio_to_pptx(source, out, page=...)`; MCP: `diagram_build` /
`diagram_install_skill`.

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
failure will not corrupt an existing `.pptx`. On first conversion a `.audited.html`
working copy is created next to the source HTML; later audit fixes edit the copy, never
the original. **As of 0.10.2 the copy is rebuilt automatically whenever the source HTML
is newer** (edits to the source take effect immediately instead of silently reusing a
stale copy). A local `<img src>` referencing a missing file now fails the conversion
with the missing paths listed (0.10.2+; no more silent blank placeholder images). The
conversion pipeline requires `offipy[deck]` and chromium:

```bash
pip install "offipy[deck]"
playwright install chromium
```

`open_live` copies the `.pptx` to an `offipy-live-*` copy in the system temp dir before
showing it in PowerPoint — PowerPoint locks the copy, never your output file, so
`render(overwrite=True)` to the same path keeps working across iterations (#22). Close the
live presentation and free the copy with `deck.close_live(doc_id)`; if a target file is
otherwise held open, `render` raises an actionable error (suggesting `close_live` /
`offipy quit ppt`, or a different output name).

For `data-asset` / `data-primitive` / `data-asset-param-*` / `data-asset-placement`
asset declarations, the `asset://` URI, providers and `assets.json` provenance, see
[Asset System](assets.en.md).

## Python API

```python
from offipy import Excel

with Excel() as x:
    book = x.new_book()                    # "book<hex>" (high-entropy doc_id)
    x.set_cell(1, "A1", 42, doc_id=book)   # destructive ops need an explicit doc_id
    assert x.read_range(1, "A1:A1", doc_id=book) == [[42.0]]
    x.quit()
```

The Python API returns the **raw values** of App methods (`new_book` → the `"book<hex>"`
string, `read_range` → a 2D list). `OperationResult` is the **return contract of the
HTTP `/call`** (`{ok, operation, resource_id, message, data}`, HTTP-only); MCP returns
the `data` payload — the three entry points have different return shapes; see the
honest comparison in [api.en.md](api.en.md).

PPTX quality audit needs no Office and never opens PowerPoint:

```python
from offipy import audit_pptx, compare_pptx, Severity

report = audit_pptx("deck.pptx")
if report.max_severity is not None and report.max_severity >= Severity.HIGH:
    print("HIGH issues found; refusing to ship")
print(report.to_markdown())

diff = compare_pptx("baseline.pptx", "candidate.pptx")
if diff.gate_severity() is not None and diff.gate_severity() >= Severity.MID:
    print("candidate adds/worsens MID+ issues vs baseline")
```

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
