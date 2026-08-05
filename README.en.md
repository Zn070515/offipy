> [中文](README.md)

# offipy

Live Microsoft Office automation via COM (session-based) + an HTML-first editable PPTX pipeline.
Goal: let Claude independently produce **polished, aesthetically sound, substantive** Office deliverables (Word / PPT / Excel).

- **Library / command**: `pip install offipy`, `import offipy`, CLI command `offipy`
- **Current version**: 0.9.0a1 (TestPyPI pre-release; initial stable release will be 1.0.0)

## Features

- **Session-based resident server**: keeps Office windows alive across calls, with document / workbook / presentation state preserved
- **Atomic operations for the three suites**: Word / Excel / PowerPoint add-edit-delete + save / export PDF
  - Excel also includes formatting capabilities — merged cells, borders, conditional formatting (cell rules / data bars / color scales), frozen panes, print setup, row height / number format / auto column width
  - Word also includes layout capabilities — a style system (character / paragraph formatting), page structure (header/footer, page numbers, page setup, table of contents), lists and tables (merge / borders / column width / row height / auto-fit), document helpers (find & replace / images / page breaks)
- **Real-time document session semantics**: ops default to acting on the user's **currently active** document (ActiveDocument / ActiveWorkbook / ActivePresentation), and are automatically re-established after a window closes
- **Disconnect self-healing**: automatically rebuilds the session when the user closes a window or Office exits
- **HTML-first pipeline + design system**: Claude writes HTML slides → natively editable `.pptx` → live presentation + visual iteration; built-in design tokens, 3 themes, 11 layouts, aesthetic audit, automatic pick, feedback learning (see "Design system" below)
- **MCP server**: exposes all three-suite operations as MCP tools, so Claude Desktop and similar can drive real Office directly
- **Environment diagnostics**: `offipy check` one-shot check of Python / dependencies / the Office three-suite / browser / server readiness (`--json` machine-readable, non-zero exit code on failure)
- **Server process management**: `offipy server status|stop|restart` uses a real `/status` handshake plus PID file / netstat probing to manage the resident process
- **Agent read-back (read-only)**: `word read_doc_text` / `ppt read_slide_texts` / `excel read_range`
  read the document's text layer back (for the agent to iterate on), exposed via CLI / RPC / MCP
- **High-level API**: `offipy.Excel() / Word() / Ppt()` context managers, driving the library directly (see "Python API" below)

## Requirements

- Windows + Microsoft Office installed (Word / Excel / PowerPoint)
- Python ≥ 3.10 (this repository's dev environment is 3.12)
- `import offipy` works on non-Windows; calling Office APIs raises `UnsupportedPlatformError`
- See [`docs/compatibility.en.md`](docs/compatibility.en.md) for the supported platform / Office / Python compatibility matrix
  (Tested / Expected / Unsupported columns)

## Installation

```bash
uv venv --python 3.12 .venv
uv pip install -e ".[all]"            # everything (office COM + deck pipeline + MCP)
uv run playwright install chromium    # the converter needs chromium for DOM measurement
```

The core `import offipy` has zero extra dependencies; install extras incrementally by use case:

- `.[office]`: Windows COM automation (Word/Excel/PowerPoint)
- `.[deck]`: HTML→editable PPTX deck pipeline (python-pptx / lxml / fonttools / playwright / Pillow)
- `.[mcp]`: MCP server (`offipy mcp`, for Claude Desktop and similar)
- `.[all]`: all of the above

The converter itself is vendored into the wheel, so it works right after install; the deck pipeline additionally needs `playwright install chromium`.

## Session semantics (read me)

`offipy` is **session-based**, not a one-shot script:

- The first call automatically starts a resident server in the background (`127.0.0.1:8890`); all subsequent operations go to the same process.
- Ops act on the **user's currently active** document — whichever workbook you have active in Excel is the one `set_cell` writes to.
  **No implicit creation**: when no document is open, ops that need a document raise `TargetNotFoundError`, telling you to run
  `new_book` / `open_book` first. Use `get_target` to query the current active target identity:
  `offipy excel get_target` → `{"app": "excel", "doc_id": "book1", "name": "Book1", "path": "..."}`
  (`null` if none). `doc_id` is a stable session identifier that stays valid across calls and does not change when the document is renamed.
- `activate(doc_id)` sets the given document as the active target and **syncs the real UI** (Excel `Workbook.Activate()`,
  Word `Document.Activate()`, PPT activates the window containing that document); on sync failure it rolls back and raises `ComOperationError`;
  `list_docs` truthfully returns the registered handles as `{doc_id: {"name", "path", "active"}}` (it does not implicitly enumerate unregistered ones).
- Destructive ops can carry an `expected_target` (`{"doc_id": ...}` / `{"name": ...}` / `{"path": ...}`, combinable)
  for **target binding**: resolve-once — the server first resolves the target doc_id from the binding keys, then executes the operation with the resolved result;
  a binding failure raises `TargetNotFoundError` — eliminating "validate A, execute B" and preventing accidental modification after switching to another document
  (available on the Python client / RPC layer, bypassing focus-based routing).
- Commands like `quit excel` close the app; `__exit__` (Python API) does **not** close the Office window, so windows and documents stay alive across calls.
- After the user manually closes a window, the next call automatically rebuilds the session (disconnect self-healing).

## Security model

- The server listens only on `127.0.0.1` and is not exposed to the public network.
- **Bearer token authentication**: a random token is generated at startup (env var `OFFIPY_SERVER_TOKEN` takes priority, otherwise persisted to
  `%LOCALAPPDATA%\offipy\token`). Every request needs `Authorization: Bearer <token>`, otherwise 401.
- Op allowlist: only the public methods of each app class are exposed; request body limit 16MB; `Content-Type` must be JSON;
  `POST` accepts only `/call` and the authenticated `/shutdown`, all other paths return 404; response body limit 64MB.
- **Process ownership (no-kill policy)**: `server status` is **read-only** — if not running it only reports "not running", it does **not implicitly start** the server;
  `server stop` prefers the authenticated `/shutdown` for graceful shutdown; when the token mismatches (`auth_fail`) or the port is occupied by a non-offipy
  process that cannot be proven to belong to us (`mismatch`), it **never kills** anything and only suggests manual handling.
- **Don't leak the token to untrusted processes** — whoever has the token has read/write access to your current Office session.
- See [`SECURITY.en.md`](SECURITY.en.md).

## Usage

```bash
# First call automatically starts a resident server in the background; all later operations hit the same process
offipy excel new_book
offipy excel set_cell --sheet 1 --cell A1 --value 100
offipy excel format_cell --sheet 1 --cell A1 --bold true --size 14 --bg "#38BDF8"
offipy excel merge_cells --sheet 1 --range_addr A1:B2
offipy excel set_border --sheet 1 --range_addr A1:D5 --side all --style continuous --weight thin --color "#D0D7DE"
offipy excel add_conditional_format --sheet 1 --range_addr C2:C5 --rule cell --operator greater --value 0 --bg "#C6EFCE"
offipy excel freeze_panes --sheet 1 --rows 1 --cols 0
offipy excel page_setup --sheet 1 --orientation landscape --fit_to_pages_wide 1
offipy excel set_number_format --sheet 1 --range_addr B2:B5 --fmt "#,##0"
offipy excel autofit --sheet 1 --range_addr A1:D5 --rows false

offipy word new_doc
offipy word write_line --text "你好，世界"
offipy word format_text --paragraph 1 --bold true --size 18 --color "#2251FF"
offipy word format_paragraph --paragraph 1 --alignment center --line_spacing double
offipy word set_header_text --text "季度报告"
offipy word add_page_number --alignment center
offipy word page_setup --orientation landscape --paper a4 --top_margin 60
offipy word insert_toc --levels 3
offipy word add_list --style bullet
offipy word merge_table_cells --table_idx 1 --start_row 1 --start_col 1 --end_row 1 --end_col 3
offipy word set_table_border --table_idx 1 --style single --color "#9AA5B1" --sides all
offipy word set_table_col_width --table_idx 1 --col 1 --width 140
offipy word find_replace --find 季度 --replace 半年度 --replace_all true
offipy word insert_image --path out/cover.png --width 360
offipy word insert_page_break

offipy ppt new_pres
offipy ppt add_slide --layout 2
offipy ppt set_title --slide_idx 1 --text "标题"

offipy check            # environment readiness diagnostics: Python/deps/Office/browser/server (--json machine-readable)
offipy server status    # resident server status (/status handshake, read-only, doesn't start it); stop / restart likewise
offipy excel get_target # query current active target identity {app,doc_id,name,path} (null if none)
offipy word read_doc_text            # agent read-back: full document text
offipy ppt read_slide_texts          # agent read-back: per-slide title/body/notes
offipy excel read_range --sheet 1 --range_addr A1:B2   # agent read-back: 2D range values
offipy quit excel
```

Complex parameters are passed through with `--payload '<json>'` (overriding same-named kwargs); repeating `--key` aggregates into a list.

## Python API

```python
from offipy import Excel, Word, Ppt

with Excel() as x:  # context manager; __exit__ does not close the Office window (session semantics)
    x.new_book()
    x.set_cell(1, "A1", 100)
    x.save("out/report.xlsx")

with Ppt() as p:
    p.new_pres()
    p.add_slide(2)
    p.set_title(1, "标题")
```

Ops not explicitly defined are proxied to the underlying app via `__getattr__`; offipy exceptions (the `OffipyError` family) pass through unchanged.

**Return contract**: the three entry points (Python API / HTTP RPC / MCP) share the same source but have **different return shapes**; see
[`docs/api.en.md`](docs/api.en.md) for the honest comparison:
- **Python API** returns the App method's raw value (`new_book`→`doc_id` string, `get_cell`→cell value,
  `get_target`→dict …); void ops return `None`; failures raise `OffipyError` domain exceptions.
- **HTTP RPC `/call`** returns an `OperationResult` (**HTTP-only contract**):
  `{ok, operation, resource_id, message, data}` (with a `result` compatibility alias). `operation` is a
  `"excel.set_cell"`-style full name; `resource_id`, e.g. `excel:book:book1`, identifies the document the operation acted on
  (`doc_id` is the stable session identifier, not the user-renamable name); `data` is the operation result (raw values for
  read ops, `null` for void ops). Raw COM objects never leak out.
- **MCP tools** return the operation `data` payload (raw values for read ops; `"ok (<op>)"` for void ops).

**Exception contract (error_code)**: failures are uniformly `OffipyError` subclasses, each carrying a `code`:
`InvalidArgumentError` (`invalid_argument`) / `TargetNotFoundError` (`target_not_found`) /
`FileConflictError` (`file_conflict`) / `ComOperationError` (`com_operation`, preserves `hresult`) /
`ProtocolError` (`protocol`). The RPC failure response's `error_code` maps one-to-one to exceptions; the client maps back to the
corresponding domain exception, so the three entry points share the same source.

## Design system (HTML deck)

When Claude writes deck HTML, it only needs to reference design tokens (CSS variables) and tag each slide with `data-layout`; themes and layouts are injected at render time, so one HTML file reskins by switching themes. See [`examples/decks/design-system/deck.html`](examples/decks/design-system/deck.html) for an example.

```html
<head>
  <style data-theme="mckinsey"></style>   <!-- theme placeholder: replaced by render(theme=) -->
  <style data-layouts></style>            <!-- layout placeholder: replaced by render(apply_layouts=) -->
</head>
<section class="slide hero" data-pptx-slide data-layout="hero-title">…</section>
```

- **Built-in themes**: `mckinsey` (consulting blue) / `academic` (academic minimalism) / `dark-tech` (dark tech) — one token set covers all, via `design.theme_css()` / `design.inject_theme()`
- **Named layout library**: `hero-title` / `split-2col` / `cards-3` / `big-number` / `quote-frame` / `timeline` / `comparison` / `chart-dominant` / `icons-row` / `portrait-feature` / `closer` — injected by reference via `layouts.inject_layouts()`
- **Aesthetic audit**: measurements produced by conversion → scored on whitespace ratio / number of font-size levels / colors per slide / contrast / cross-slide consistency; `aesthetic.audit()` outputs a scored report for iteration
- **Automatic pick**: `autopick.pick()` recommends a theme + per-slide layout + reasoning from the content structure (pure rules, overridable)
- **Content workflow**: standardizes the skeleton of "substance" — write a markdown outline
  (`# Title` + each `## section` = one slide, `- ` bullet points, body text, `@layout:` directives),
  `offipy deck outline --input outline.md --out deck.html` produces an HTML skeleton in one step with
  layouts auto-inferred, then `offipy deck make --theme <theme> --layouts` injects theme and layouts to finalize.
  See [`examples/outline/quarterly-review.md`](examples/outline/quarterly-review.md) for an example.

  ```bash
  offipy deck outline --input examples/outline/quarterly-review.md --out out/quarterly.html
  offipy deck make --html out/quarterly.html --theme mckinsey --layouts --out out/quarterly.pptx
  ```
- **Native charts**: tag a chart area with `data-chart="<type>"` (`bar` / `line` / `pie`), put the data in a container
  `data-chart-data` JSON attribute or an in-page `<script type="application/json" data-chart-target="<selector>">` block,
  and `deck make --layouts` replaces them with PowerPoint-native editable charts (double-click to edit data), no longer pasted images.
  **Prerequisite**: the slide containing a chart must reference the chart-dominant layout (`data-layout="chart-dominant"`) so the container has a measurable
  surface rectangle. See [`examples/decks/charts/chart-demo.html`](examples/decks/charts/chart-demo.html) for an example.

  ```html
  <div class="chart" data-chart="bar"
       data-chart-data='{"categories":["Q1","Q2"],"series":[{"name":"营收","values":[40,70]}]}'></div>
  ```

  In the outline, use `@chart: <type>` + `@chart-data: <JSON>` to declare chart slides; the skeleton automatically lands on the chart-dominant layout.
- **Native icons**: tag an icon container with an empty `<svg data-icon="<set>:<name>" viewBox=... width=.. height=..>`,
  and `deck make --layouts` replaces it with a PowerPoint-native freeform vector icon (double-click shows editable
  anchor points, not an image). Two built-in sets — Phosphor (`ph:`, 256 viewBox, filled) + Lucide (`lu:`, 24 viewBox,
  line style) (vendored in `src/offipy/assets/icons/`, updated via `scripts/fetch_icons.py`);
  icon color inherits the container's `color` (setting `color: var(--accent)` in HTML applies the theme color; if the container has no
  color set, it defaults to the current theme's `--accent`). Lucide line icons render with round line caps/corners (aligned with the
  source SVG design, not images or pasted graphics). See
  [`examples/decks/icons/icons-demo.html`](examples/decks/icons/icons-demo.html) for an example.

  ```html
  <svg class="icon" data-icon="ph:check-circle" viewBox="0 0 256 256" width="72" height="72"></svg>
  ```

  In the outline, use `@icons: <name>[:label]; ...` to declare icon rows (default `ph:` prefix); the skeleton automatically lands on the
  `icons-row` layout (more than 3 icons auto-wrap into 3 columns per row). See
  [`examples/outline/icons-demo.md`](examples/outline/icons-demo.md) for an example.
- **Feedback learning**: post-audit dispositions (fixed / accepted / ignored) are recorded to `~/.offipy/feedback.jsonl`; `feedback.dimension_weights()` reweights the audit weights, getting stricter the more it fixes (P2 validation build)

```python
from offipy import deck, aesthetic

pptx = deck.render("deck.html", theme="mckinsey", apply_layouts=True)  # inject theme + layouts
report = aesthetic.audit("deck.html")  # read measurement and score
print(report.markdown())
```

## MCP server (Claude integration)

`offipy mcp` starts an MCP stdio server that exposes all Word / Excel / PowerPoint operations as MCP tools
(`ppt_set_title`, `word_write_line`, `excel_set_cell`, etc.). Tool calls are equivalent to `offipy` commands,
acting on the user's currently active document / workbook / presentation, with windows visible in real time.

Add it to Claude Desktop's `claude_desktop_config.json`. `<OFFIPY_ROOT>` is the absolute path to your local repository (Windows example: `C:\\path\\to\\offipy`); **do not commit your real machine path**:

```json
{
  "mcpServers": {
    "offipy": {
      "command": "<OFFIPY_ROOT>\\.venv\\Scripts\\python.exe",
      "args": ["-m", "offipy.mcp_server"]
    }
  }
}
```

Manual verification (no Office needed, handshake only):

```bash
offipy mcp        # blocks, waiting for a stdio client to connect
```

## Release hardening (ChatGPT review fix mapping)

0.9.0a1 closes out all fixes from the third-party review, item by item:

| Review item | Fix |
|--------|------|
| P0-1 converter not in wheel | converter vendored into `src/offipy/_vendor/`, so the deck pipeline works after `pip install` |
| P0-2 missing chromium precheck | checks the browser before rendering; raises `ConversionError` with an install hint if missing |
| P0-3 library layer raises SystemExit | new offipy exception hierarchy; library layer no longer calls `sys.exit` |
| P0-4 server security | Bearer token + `/status` + 16MB limit + Content-Type + op allowlist |
| P0-5 mcp dependency too broad | narrowed to `mcp>=2.0,<3.0` |
| P0-6 no release gate | release workflow completed: lint→format→mypy→pytest→tag version check→build→twine→install smoke→`--verify-tag` |
| P0-7 cannot import on non-Windows | lazy COM import, cross-platform compatible |
| P1 all | CLI renamed `offipy` + complex args, high-level API, real-time document session semantics, save overwrite protection, metadata/docs/types/CI governance |
| P2 architecture | the three architectural items landed: session semantics, CLI argument system, high-level API |

**round-3 (ChatGPT_v3 review)**: four main threads — target identity / process ownership / failure atomicity / entry-point consistency, closed out batch by batch over 11 batches:

| Batch | Fix |
|------|------|
| Process ownership + HTTP boundaries | `_probe` four states (ok/auth_fail/mismatch/down), `ensure_server` no-kill policy, authenticated `/shutdown`, `server status` read-only, path 404 / negative length 400 / response limit 500 |
| Target identity + read-only, no create | `get_target`, `/status` targets, `expected_target` binding, raise `TargetNotFoundError` when no document |
| DisplayAlerts scope + deck atomicity + CLI real booleans | save/restore within ops, temp file + `os.replace` atomic replacement, `store_true` booleans |
| Threads + worker + returns + exceptions | `ThreadingHTTPServer` + single COM worker queue, `OperationResult` contract, domain exception hierarchy A |
| schema single source of truth | new RPC changes only touch `schema.py`; server/CLI/MCP three entry points derive from it |
| extras + branding | `office/deck/mcp/all` split, `py.typed`, unified `offipy` |
| CI matrix + non-Windows | pure-module (coverage gate) / windows 3.10–3.13 / wheel-smoke / office-real; MCP in-memory protocol-layer tests |

**round-4 (ChatGPT_v4 review)**: explicit target semantics + bounded resources + release gates + tightened conversion boundaries:

| Main thread | Fix |
|------|------|
| Explicit target semantics (P0-4/5/6) | `doc_id` is authoritative; `expected_target` is **resolve-once** (three keys `doc_id`/`name`/`path`, rejects empty objects / unknown keys, injects the resolved doc_id into args); `activate()` syncs the real UI (rolls back and raises `ComOperationError` on failure); `resource_id` uses doc_id; `list_docs` only reports registered handles (including `active`); `get_target(doc_id=)` for explicit queries |
| Bounded resources + idempotency (§4/5) | client timeout 600s aligned with server; `request_id` idempotency cache (repeated retries don't re-execute); COM queue limit 64 + concurrent-thread limit 16, 503 when full; PID file includes `port/pid/token_sha256/started_at` for ownership verification; token file 0o600 |
| Deck atomic rendering (§7) | `render` temp files switch to `mkstemp` (random names in the same directory, no collision under concurrency); temp files cleaned up on failure, never destroys an existing .pptx; missing source HTML raises `InvalidArgumentError` |
| Release gates (P0-2/3) | `release.yml` gate chain quality → office-real (must pass on real machine) → gh-release → publish-testpypi → publish-pypi (OIDC Trusted Publishing); `ci.yml` office-real becomes the PR merge gate; `docs/release.md` release manual |
| Tightened conversion boundaries (§6/11/12) | `_parse_cell` tightened (rejects out-of-range XFD/1048576); `set_title`/`set_body` explicitly error on empty input; CLI `--port` subcommand SUPPRESS inheritance; `/call` non-dict args → 400; `offipy check --profile`; `install_smoke.py --profile`; license `MIT AND ISC`; dependabot |
| save/close prevent dialog pop-ups | `save()`/`close_book()`/`close_doc()` auto-save never-saved documents to the same directory and return the absolute path, without the Save As dialog; `overwrite` overwrite protection fail-fast |

## Development

```bash
uv sync --extra dev                     # install dev deps (ruff / mypy / pytest)
uv run ruff check .                     # lint
uv run ruff format --check .            # format
uv run mypy src/offipy                  # types
uv run pytest tests -q                  # tests (COM integration tests skip automatically without Office)
```

See [`CONTRIBUTING.en.md`](CONTRIBUTING.en.md) for contribution guidelines and gates.

## Structure

```
src/offipy/
  core.py       # COM lifecycle/session management + active_doc/doc_alive session semantics
  exceptions.py # offipy exception hierarchy (OffipyError + 10 subclasses, strategy-A domain exceptions)
  schema.py     # operation schema single source of truth (OpSpec table; server/CLI/MCP three entry points derive from it)
  result.py     # OperationResult return contract (HTTP-only: ok/operation/resource_id/message/data)
  paths.py      # user data directory / converter data directory / ensure_writable overwrite protection
  server.py     # resident session HTTP server (token auth + worker queue + /status + /shutdown)
  cli.py        # `offipy` command entry point (complex args: repeated flag → list / --payload JSON)
  api.py        # high-level API facade: Excel() / Word() / Ppt() context managers
  mcp_server.py # MCP stdio server (three-suite operations → MCP tools)
  excel.py / word.py / ppt.py   # three-suite atomic operations
  client.py     # HTTP client for the server (reused by the HTML pipeline) *
  envcheck.py   # `offipy check` environment readiness diagnostics (grouped checklist + --json) *
  deck.py       # HTML → editable PPTX pipeline (render/open_live/export_slides) *
  design.py     # design system: token model + 3 built-in themes + .slide base styles *
  layouts.py    # named layout library: 11 layout components + data-layout injection *
  charts.py     # native charts: chart declaration parsing + native editable chart injection (bar/line/pie) *
  icons.py      # native icons: SVG path flattening + freeform vector icon injection (ph/lu dual sets) *
  assets/icons/ # vendored icon assets (Phosphor fill + Lucide) + manifest + LICENSE *
  aesthetic.py  # aesthetic audit: whitespace/font-size hierarchy/color count/contrast/consistency → scored report *
  autopick.py   # automatic pick: content structure → recommended theme + per-slide layout + reasoning *
  feedback.py   # feedback learning: audit dispositions → dimension weights (P2 validation build) *
  outline.py    # content workflow: markdown outline → per-slide structured content → HTML skeleton *
  _vendor/html_to_editable_pptx/  # vendored HTML→PPTX converter (third-party code, MIT) *
tests/        # pytest
docs/         # protocol (protocol.md) / return contract (api.md) / deprecation (deprecation.md) / compatibility matrix (compatibility.md) / release manual (release.md)
examples/     # runnable examples (decks / outline / excel / word)
```

\* Added by the "HTML-first editable PPTX pipeline" and "M1/M2/M3 content workflow" plans.

## License

MIT AND ISC (SPDX expression) — offipy itself and the vendored converter / Phosphor icons are MIT,
the embedded Lucide icons are ISC. The license texts are distributed with the artifacts: the root `LICENSE` +
[`THIRD_PARTY_NOTICES.en.md`](THIRD_PARTY_NOTICES.en.md); each icon set's `LICENSE-*.txt` is in
`src/offipy/assets/icons/`.

## Feedback and issues

- **Bugs / feature requests**: file them at [GitHub Issues](https://github.com/Zn070515/office-kit/issues),
  attaching `offipy check --json` output and a minimal repro.
- **Pre-release versions**: smoke-test TestPyPI with `scripts/pypi_smoke.py --index https://test.pypi.org/simple --version <pre-release version>`
  (see "Pre-release numbering policy" in the CHANGELOG). Pre-release versions are for validation only; before the initial stable release,
  the top three — `__version__` / tag / CHANGELOG — must stay consistent.
