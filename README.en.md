> [中文](README.md)

# offipy

Live Microsoft Office automation via COM (session-based) + an HTML-first editable PPTX pipeline.
Built for Python developers and AI agents to independently produce **polished, aesthetically sound, substantive** Office deliverables (Word / PPT / Excel).

- **Library / command**: `pip install offipy`, `import offipy`, CLI command `offipy`
- **Current version**: 0.19.0 (the current stable release; 1.0.0 will follow broader API validation)

## Features

- **Session-based resident server**: keeps Office windows alive across calls, with document / workbook / presentation state preserved
- **Atomic operations for the three suites**: Word / Excel / PowerPoint add-edit-delete + save / export PDF
  - Excel also includes formatting capabilities — merged cells, borders, conditional formatting (cell rules / data bars / color scales), frozen panes, print setup, row height / number format / auto column width
  - Word also includes layout capabilities — a style system (character / paragraph formatting), page structure (header/footer, page numbers, page setup, table of contents), lists and tables (merge / borders / column width / row height / auto-fit), document helpers (find & replace / images / page breaks)
- **Real-time document session semantics**: read ops default to acting on the user's **currently active** document (ActiveDocument / ActiveWorkbook / ActivePresentation); destructive ops require an explicit `doc_id`, `follow_active=True`, or an `expected_target` binding, so they can never silently modify the wrong document; sessions auto-rebuild after a window closes
- **Disconnect self-healing**: automatically rebuilds the session when the user closes a window or Office exits
- **HTML-first pipeline + design system**: Claude writes HTML slides → natively editable `.pptx` → live presentation + visual iteration; built-in design tokens, 3 themes, 11 layouts, aesthetic audit, automatic pick, feedback learning (see "Design system" below)
- **MCP server**: exposes all three-suite operations as MCP tools, so Claude Desktop and similar can drive real Office directly
- **Environment diagnostics**: `offipy check` one-shot check of Python / dependencies / the Office three-suite / browser / server readiness (`--json` machine-readable, non-zero exit code on failure)
- **Server process management**: `offipy server status|stop|restart` uses a real `/status` handshake plus PID file / netstat probing to manage the resident process
- **Agent read-back (read-only)**: `word read_doc_text` / `ppt read_slide_summary` (per-slide title/body/notes) `ppt read_slide_texts --slide_idx N` (per-slide per-shape text) / `excel read_range`
  read the document's text layer back (for the agent to iterate on), exposed via CLI / RPC / MCP
- **PPTX static quality gate**: `offipy audit` (see [`docs/audit.en.md`](https://github.com/Zn070515/offipy/blob/main/docs/audit.en.md)) — no PowerPoint, no
  Microsoft Office; parses `.pptx` directly to check out-of-bounds / edge-adjacent / overlap / text overflow /
  autofit risks, emits text / json / markdown / html reports, and blocks non-conforming artifacts by severity
  threshold (`--fail-on`); `compare_pptx` baseline regression (see
  [`docs/audit-baseline.en.md`](https://github.com/Zn070515/offipy/blob/main/docs/audit-baseline.en.md)) only blocks candidate **added/worsened** issues;
  `deck render_with_report` turns HTML→PPTX generation into a gate
- **Art analysis (offipy.art)**: pure-stdlib, deterministic, **advisory-only** visual/typographic quality analysis (see
  [`docs/art.en.md`](https://github.com/Zn070515/offipy/blob/main/docs/art.en.md)) — `build_scene` abstracts a slide
  into an ArtScene, 5 dimension rule sets (hierarchy / composition / typography / color / media) with grade /
  confidence / evidence_coverage separated and evidence-poor dimensions downgraded instead of false-reporting;
  `analyze_deck` evaluates all three sources (measurements + pptx geometry + slides_dir per-page PNG
  pixels, lazy Pillow) in one call, and
  `deck.render_with_quality_report` turns generation into a quality reference
- **Learnable feedback (v0.18)**: numpy MLP + registry-based inputs/outputs (FEATURES / OUTPUTS),
  `offipy feedback train` / `status` / `append` / `recommend` / `apply` / `reschema`, cold-start falls
  back to v2, core stays numpy-free. Learning quality (#115-#122): standardized preprocessing (zero-variance
  drop + high-correlation dedup + z-score), ensemble K=5 + calibration + abstain + OOD, sample-level
  repeated stratified CV, capacity-adaptive warnings, per-rule sample diagnosis; consumer side
  (#130-#160): status reports stale/corrupt, `experimental_score_mode` annotates the score source,
  `quality_score_coverage` reports confident-subset coverage, severity_shift has provenance
  (override + details), rule.delta is visible in the report, and `recommend`/`apply` persist to
  profile storage — `deck audit --feedback-dir` passes through experimental score
- **Shape-level PPT reading and editing**: `ppt read_shapes` reads the shape tree (frozen `ShapeInfo` /
  shape-type contract), with `set_shape_geometry/text/font/fill/outline/visible`, `delete_shape`, and
  `set_shape_z_order` to edit real PowerPoint objects incrementally
- **CLI error contract**: `InvalidArgumentError → exit 2` (usage / argument / pre-runtime invalid input),
  `OffipyError → exit 1` (runtime domain failure), stderr without leaked tracebacks; `audit` keeps its
  0/1/2/3 and `deck audit` keeps its 0/1 dedicated contracts
- **High-level API**: `offipy.Excel() / Word() / Ppt()` context managers, driving the library directly (see "Python API" below)
- **diagram-design skill integration (Agent-native)**: `offipy diagram build` turns Mermaid / draw.io
  diagrams that a host agent designed per the artifact contract into editable PPTX;
  `offipy diagram install_skill` installs the diagram-design guide plus the offipy-diagram
  contract-bridge skill into the host agent's skill directory

## Requirements

- Windows + Microsoft Office installed (Word / Excel / PowerPoint)
- Python ≥ 3.10 (this repository's dev environment is 3.12)
- `import offipy` works on non-Windows; calling Office APIs raises `UnsupportedPlatformError`
- See [`docs/compatibility.en.md`](https://github.com/Zn070515/offipy/blob/main/docs/compatibility.en.md) for the supported platform / Office / Python compatibility matrix
  (Tested / Expected / Unsupported columns)

## Installation

```bash
py -m pip install "offipy[all]"       # everything (office COM + deck pipeline + MCP)
py -m playwright install chromium     # the converter needs chromium for DOM measurement
offipy check --profile all            # one-shot environment readiness check (Python/deps/Office/browser/server)
```

The core `import offipy` has zero extra dependencies; install extras incrementally by use case:

- `offipy[office]`: Windows COM automation (Word/Excel/PowerPoint)
- `offipy[deck]`: HTML→editable PPTX deck pipeline (python-pptx / lxml / fonttools / playwright / Pillow)
- `offipy[mcp]`: MCP server (`offipy mcp`, for Claude Desktop and similar)
- `offipy[all]`: all of the above

The converter itself is vendored into the wheel, so it works right after install; the deck pipeline additionally needs `playwright install chromium`.

## Session semantics (read me)

`offipy` is **session-based**, not a one-shot script:

- The first call automatically starts a resident server in the background (`127.0.0.1:8890`); all subsequent operations go to the same process.
- **`doc_id` is the authoritative target identifier**: `new_book` / `new_doc` / `new_pres` and the `open_*` ops return a `doc_id`
  (`book<hex>` / `doc<hex>` / `pres<hex>`, high-entropy and opaque), which is stable within the session, stays valid across calls,
  and does not change when the document is renamed. Use `get_target` to query the current active target identity:
  `offipy excel get_target` → `{"app": "excel", "doc_id": "book<hex>", "name": "Book1", "path": "..."}`
  (`null` if none).
- **Read ops** (`get_cell` / `read_range` / `read_doc_text` / `read_slide_summary` / `read_slide_texts` / `get_target` …) default to acting on the
  user's **currently active** document (ActiveDocument / ActiveWorkbook / ActivePresentation), resolved in real time — never from a stale
  internal cache. When no document is open, they raise `TargetNotFoundError`, telling you to run `new_*` / `open_*` first.
- **Destructive ops** (writes / formatting / save / close, etc.) **refuse to run by default**; you must provide one of three:
  - an explicit `doc_id=<the session-returned identifier>` (CLI `--doc-id`);
  - or `follow_active=True` — an explicit declaration to "follow the currently active document" (CLI `--follow-active`);
  - or an `expected_target` binding (below).
  If none is present, they raise `InvalidArgumentError` prompting you for a target — **they never silently write to the currently active document**.
- **`expected_target` binding** (available across CLI `--expected-target '<json>'` / MCP tool arguments / the Remote client,
  all three entry points): `{"doc_id": ...}` / `{"name": ...}` / `{"path": ...}`, combinable, for **target binding**:
  resolve-once — the server first resolves the target doc_id from the binding keys, then executes the operation with the resolved result;
  a binding failure raises `TargetNotFoundError` — eliminating "validate A, execute B" and preventing accidental modification after switching to another document
  (bypassing focus-based routing).
- `activate(doc_id)` sets the given document as the active target and **syncs the real UI** (Excel `Workbook.Activate()`,
  Word `Document.Activate()`, PPT activates the window containing that document); on sync failure it rolls back and raises `ComOperationError`;
  `list_docs` truthfully returns the registered handles as `{doc_id: {"name", "path", "active"}}` (it does not implicitly enumerate unregistered ones).
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
# Destructive ops need a target: --doc-id <the session-returned id> / --follow-active (follow the active document) /
# --expected-target '<json>' (doc_id/name/path binding). The examples use --follow-active to follow the newly created/opened document.
offipy excel new_book
offipy excel set_cell --sheet 1 --cell A1 --value 100 --follow-active
offipy excel format_cell --sheet 1 --cell A1 --bold true --size 14 --bg "#38BDF8" --follow-active
offipy excel merge_cells --sheet 1 --range_addr A1:B2 --follow-active
offipy excel set_border --sheet 1 --range_addr A1:D5 --side all --style continuous --weight thin --color "#D0D7DE" --follow-active
offipy excel add_conditional_format --sheet 1 --range_addr C2:C5 --rule cell --operator greater --value 0 --bg "#C6EFCE" --follow-active
offipy excel freeze_panes --sheet 1 --rows 1 --cols 0 --follow-active
offipy excel page_setup --sheet 1 --orientation landscape --fit_to_pages_wide 1 --follow-active
offipy excel set_number_format --sheet 1 --range_addr B2:B5 --fmt "#,##0" --follow-active
offipy excel autofit --sheet 1 --range_addr A1:D5 --rows false --follow-active

offipy word new_doc
offipy word write_line --text "你好，世界" --follow-active
offipy word format_text --paragraph 1 --bold true --size 18 --color "#2251FF" --follow-active
offipy word format_paragraph --paragraph 1 --alignment center --line_spacing double --follow-active
offipy word set_header_text --text "季度报告" --follow-active
offipy word add_page_number --alignment center --follow-active
offipy word page_setup --orientation landscape --paper a4 --top_margin 60 --follow-active
offipy word insert_toc --levels 3 --follow-active
offipy word add_list --lines "first item" --lines "second item" --style bullet --follow-active
offipy word merge_table_cells --table_idx 1 --start_row 1 --start_col 1 --end_row 1 --end_col 3 --follow-active
offipy word set_table_border --table_idx 1 --style single --color "#9AA5B1" --sides all --follow-active
offipy word set_table_col_width --table_idx 1 --col 1 --width 140 --follow-active
offipy word find_replace --find 季度 --replace 半年度 --replace_all true --follow-active
offipy word insert_image --path out/cover.png --width 360 --follow-active
offipy word insert_page_break --follow-active

offipy ppt new_pres
offipy ppt add_slide --layout 2 --follow-active
offipy ppt set_title --slide_idx 1 --text "标题" --follow-active

offipy check            # environment readiness diagnostics: Python/deps/Office/browser/server (--json machine-readable)
offipy server status    # resident server status (/status handshake, read-only, doesn't start it); stop / restart likewise
offipy excel get_target # query current active target identity {app,doc_id,name,path} (null if none)
offipy word read_doc_text            # agent read-back: full document text
offipy ppt read_slide_summary        # agent read-back: per-slide title/body/notes summary
offipy ppt read_slide_texts --slide_idx 1   # agent read-back: per-shape text on one slide (v0.10)
offipy excel read_range --sheet 1 --range_addr A1:B2   # agent read-back: 2D range values
offipy quit excel

offipy audit deck.pptx                    # PPTX static quality audit (text report)
offipy audit deck.pptx --fail-on HIGH     # exit code 1 at HIGH (CI gate)
offipy audit deck.pptx --format html --out audit.html --slides-dir export/  # SVG canvas report
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID     # baseline regression: only block added/worsened
offipy deck make --html deck.html --out deck.pptx --no-open \
  --audit-mode strict --fail-on HIGH --audit-report deck.audit.json        # HTML→PPTX generate-and-gate
```

Complex parameters are passed through with `--payload '<json>'` (overriding same-named kwargs); repeating `--key` aggregates into a list.
Read ops (`get_cell` / `read_range` / `read_doc_text` / `read_slide_summary` / `read_slide_texts` / `get_target`) need no target parameter.

## Python API

```python
from offipy import Excel, Word, Ppt

with Excel() as x:  # local direct COM (= offipy.direct.*), independent doc_id/thread
    doc_id = x.new_book()
    x.set_cell(1, "A1", 100, doc_id=doc_id)  # destructive ops need an explicit doc_id
    x.save("out/report.xlsx", doc_id=doc_id)

with Ppt() as p:
    pres_id = p.new_pres()
    p.add_slide(2, doc_id=pres_id)
    p.set_title(1, "标题", doc_id=pres_id)
```

**Two session models (P0-4)**:
- `Excel() / Word() / Ppt()` (equivalent to `offipy.direct.*`) — **local direct COM**,
  with a doc_id/thread/session state fully isolated from the CLI/MCP.
- `RemoteExcel() / RemoteWord() / RemotePpt()` — a **remote session** through the resident server,
  sharing the same session (same doc_id) with the CLI/MCP; the right choice when an agent needs
  "CLI/Python/tools all in one Office session":

```python
from offipy import RemoteExcel

with RemoteExcel() as x:  # connects to local 8890 by default (auto-starts the server)
    x.new_book()  # the same doc_id seen by `offipy excel list_docs`
    x.set_cell(1, "A1", 42, follow_active=True)
```

Ops not explicitly defined are proxied to the underlying app via `__getattr__`; offipy exceptions (the `OffipyError` family) pass through unchanged.

**Return contract**: the three entry points (Python API / HTTP RPC / MCP) share the same source but have **different return shapes**; see
[`docs/api.en.md`](https://github.com/Zn070515/offipy/blob/main/docs/api.en.md) for the honest comparison:
- **Python API** returns the App method's raw value (`new_book`→`doc_id` string, `get_cell`→cell value,
  `get_target`→dict …); void ops return `None`; failures raise `OffipyError` domain exceptions.
- **HTTP RPC `/call`** returns an `OperationResult` (**HTTP-only contract**):
  `{ok, operation, resource_id, message, data}` (with a `result` compatibility alias). `operation` is a
  `"excel.set_cell"`-style full name; `resource_id`, e.g. `excel:book:book<hex>`, identifies the document the operation acted on
  (`doc_id` is the stable session identifier, not the user-renamable name); `data` is the operation result (raw values for
  read ops, `null` for void ops). Raw COM objects never leak out.
- **MCP tools** return the operation `data` payload (raw values for read ops; `"ok (<op>)"` for void ops).

**Exception contract (error_code)**: failures are uniformly `OffipyError` subclasses, each carrying a `code`:
`InvalidArgumentError` (`invalid_argument`) / `TargetNotFoundError` (`target_not_found`) /
`FileConflictError` (`file_conflict`) / `ComOperationError` (`com_operation`, preserves `hresult`) /
`ProtocolError` (`protocol`). The RPC failure response's `error_code` maps one-to-one to exceptions; the client maps back to the
corresponding domain exception, so the three entry points share the same source.

## Design system (HTML deck)

When Claude writes deck HTML, it only needs to reference design tokens (CSS variables) and tag each slide with `data-layout`; themes and layouts are injected at render time, so one HTML file reskins by switching themes. See [`examples/decks/design-system/deck.html`](https://github.com/Zn070515/offipy/blob/main/examples/decks/design-system/deck.html) for an example.

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
  See [`examples/outline/quarterly-review.md`](https://github.com/Zn070515/offipy/blob/main/examples/outline/quarterly-review.md) for an example.

  ```bash
  offipy deck outline --input examples/outline/quarterly-review.md --out out/quarterly.html
  offipy deck make --html out/quarterly.html --theme mckinsey --layouts --out out/quarterly.pptx
  ```
- **Native charts**: tag a chart area with `data-chart="<type>"` (`bar` / `line` / `pie`), put the data in a container
  `data-chart-data` JSON attribute or an in-page `<script type="application/json" data-chart-target="<selector>">` block,
  and `deck make --layouts` replaces them with PowerPoint-native editable charts (double-click to edit data), no longer pasted images.
  **Prerequisite**: the slide containing a chart must reference the chart-dominant layout (`data-layout="chart-dominant"`) so the container has a measurable
  surface rectangle. See [`examples/decks/charts/chart-demo.html`](https://github.com/Zn070515/offipy/blob/main/examples/decks/charts/chart-demo.html) for an example.

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
  [`examples/decks/icons/icons-demo.html`](https://github.com/Zn070515/offipy/blob/main/examples/decks/icons/icons-demo.html) for an example.

  ```html
  <svg class="icon" data-icon="ph:check-circle" viewBox="0 0 256 256" width="72" height="72"></svg>
  ```

  In the outline, use `@icons: <name>[:label]; ...` to declare icon rows (default `ph:` prefix); the skeleton automatically lands on the
  `icons-row` layout (more than 3 icons auto-wrap into 3 columns per row). See
  [`examples/outline/icons-demo.md`](https://github.com/Zn070515/offipy/blob/main/examples/outline/icons-demo.md) for an example.
- **Native diagrams (Mermaid / draw.io)**: write `<pre class="mermaid">graph TD ...</pre>` in a page; `deck render`
  lets `offipy.diagrams` replace the Mermaid flowchart with PowerPoint-native editable shapes (reusing the
  charts injection pipeline, requires the visual-audit `measurements.json`). The standalone API
  `offipy.diagrams.mermaid_to_pptx("graph TD\nA-->B", "out.pptx")` renders a full 16:9 editable PPTX
  (TD/TB/LR/RL/BT directions, subgraph containers, Chinese labels). Note `graph`/`flowchart` must carry an
  explicit direction (bare `graph` is rejected); `%%` comments are not supported (vendored parser contract).
  draw.io is also supported: write `<div class="drawio" data-drawio="arch.drawio"></div>` in the
  HTML (relative paths are rewritten to `file://` absolute URIs in the deck pipeline, so they
  resolve even when the HTML is staged to a temp dir); `deck render` turns it into editable
  shapes, keeping the author's layout and colors, with node `fontSize` scaled to the container.
  Orthogonal/curved edges render as polylines along their waypoints (arrows kept), and
  `strokeWidth` / `rotation` / `dashPattern` propagate to the shapes. Multi-page `.drawio`
  files require `data-drawio-page="N"` (1-based) in deck injection, otherwise it errors rather
  than silently taking the first page. Standalone
  `offipy.drawio.drawio_to_pptx("arch.drawio", "out.pptx", page="架构")` renders a full 16:9
  editable PPTX (`page` accepts an int index or str page name, default first page).
- **Feedback learning**: post-audit dispositions (fixed / accepted / ignored) are recorded to `~/.offipy/feedback.jsonl`; `feedback.dimension_weights()` reweights the audit weights, getting stricter the more it fixes (P2 validation build)

```python
from offipy import deck, aesthetic

pptx = deck.render("deck.html", theme="mckinsey", apply_layouts=True)  # inject theme + layouts
report = aesthetic.audit("deck.html")  # read measurement and score
print(report.markdown())
```

## Art analysis (offipy.art)

`offipy.art` is **pure-stdlib, deterministic, advisory-only** visual/typographic quality analysis: no AI, no
Office, no python-pptx on `import offipy`. It abstracts each slide as a "scene (ArtScene)" and evaluates it
with 5 dimension rule sets (hierarchy / composition / typography / color / media). grade / confidence /
evidence_coverage are kept strictly separate — evidence-poor dimensions degrade to `insufficient_evidence`
instead of guessing false findings; the verdict is left to the caller, there is no total-score gate. Full
rules, evidence sources, and boundaries: [`docs/art.en.md`](https://github.com/Zn070515/offipy/blob/main/docs/art.en.md).

```python
from offipy import build_scene, analyze_scene, analyze_deck, render_markdown

# three sources: measurements (real browser pixels) primary, pptx geometry audit secondary,
# slides_dir per-page PNG pixels supplemental
scene = build_scene(
    measurements="out/report_audit/_cache/measurements.json",
    pptx="out/report.pptx",
    slides_dir="out/slides",  # optional: per-page PNG pixel evidence (requires Pillow, `offipy[deck]`)
)
report = analyze_scene(scene, profile="balanced")
print(render_markdown(report))

# combined entry: geometry audit + art analysis in one call
deck_report = analyze_deck(
    pptx="out/report.pptx",
    measurements="out/report_audit/_cache/measurements.json",
    profile="consulting",
)
print(deck_report.art.slides[0].by_dimension("color").status)  # assessed / insufficient_evidence
```

- **Three-source merge**: measurements provide color / font-size / text evidence, pptx provides geometry,
  and slides_dir provides per-page PNG pixel evidence (page-level background / whitespace + declared-color
  verification, lazy Pillow) — matched one-to-one by "text as strong corroboration + geometry as fallback";
  unmatched elements are kept with a warning, never silently dropped
- **Built-in profiles**: `balanced` / `consulting` / `academic` / `technology` / `event`
  (`offipy.profile_names()` to list, `get_profile(name)` to read / extend)
- **Quality-on-generate**: `deck.render_with_quality_report(html, audit_mode=..., fail_on=..., profile=...)`
  produces both the geometry audit and the art analysis after HTML→PPTX, returning a `QualityRenderResult`
  (`art_report` / `deck_quality`) — `audit_mode="strict"` only exits non-zero at the `fail_on` threshold
- **Evidence honesty**: with only `pptx=` given, `_ShapeRecord` enrichment provides font-size / font-family /
  foreground / background / opacity / fill_kind evidence (#128), so the hierarchy / typography / color
  dimensions no longer degrade from zero evidence; when pure pixel evidence (PNG / measurements) is
  missing, the dimensions that need it still degrade (`insufficient_evidence` + `art.evidence.limited`
  warning), while pure-geometry rules keep running

## MCP server (Claude integration)

`offipy mcp` starts an MCP stdio server that exposes all Word / Excel / PowerPoint operations as MCP tools
(`ppt_set_title`, `word_write_line`, `excel_set_cell`, etc.). Tool calls are equivalent to `offipy` commands,
with windows visible in real time; read ops act on the user's currently active document, and destructive tool
arguments include `expected_target` / `follow_active` (see "Session semantics" above).

Add it to Claude Desktop's `claude_desktop_config.json`. The `offipy` command must be on PATH (pip install adds it automatically); if you use a dedicated venv, point `command` at that venv's `offipy.exe` absolute path (e.g. `<venv>\\Scripts\\offipy.exe`):

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

Manual verification (no Office needed, handshake only):

```bash
offipy mcp        # blocks, waiting for a stdio client to connect
```

## Development

To develop from source (rather than the PyPI release):

```bash
uv venv --python 3.12 .venv
uv pip install -e ".[all]"            # source dev install (everything)
uv run playwright install chromium    # needed by the deck pipeline
```

```bash
uv sync --extra dev                   # install dev deps (ruff / mypy / pytest)
uv run ruff check .                   # lint
uv run ruff format --check .          # format
uv run mypy src/offipy                # types
uv run pytest tests -q                # tests (COM integration tests skip automatically without Office)
uv run mkdocs build --strict          # docs build gate (any warning fails, #108)
```

The pure-module CI gate (Linux, no Office) additionally runs `mkdocs build --strict` and
seeded-mutation fuzz of the drawio / HTML parsers (#110) — malformed input must not crash and
URL / data paths must pass the whitelist.

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
  audit/        # PPTX static quality audit & baseline regression (models/extract/geometry/roles/rules/compare/render/pptx, pure parsing, no COM)
  art/          # Art analysis (v0.12): scene models / rules / dual-source merge / render / baseline compare (pure stdlib, advisory only)
  mcp_server.py # MCP stdio server (three-suite operations → MCP tools)
  excel.py / word.py / ppt.py   # three-suite atomic operations
  client.py     # HTTP client for the server (reused by the HTML pipeline) *
  envcheck.py   # `offipy check` environment readiness diagnostics (grouped checklist + --json) *
  deck.py       # HTML → editable PPTX pipeline (render/open_live/export_slides) *
  design.py     # design system: token model + 3 built-in themes + .slide base styles *
  layouts.py    # named layout library: 11 layout components + data-layout injection *
  charts.py     # native charts: chart declaration parsing + native editable chart injection (bar/line/pie) *
  icons.py      # native icons: SVG path flattening + freeform vector icon injection (ph/lu dual sets) *
  diagrams.py   # native diagrams: Mermaid flowchart extraction + layered layout + editable shape rendering (mermaid_to_pptx / deck injection) *
  drawio.py     # native diagrams: draw.io → editable shapes (drawio_to_pptx / deck data-drawio injection) *
  diagram.py    # diagram app: Mermaid/drawio → editable PPTX (Agent-native; install_skill installs the vendored skill) *
  assets/icons/ # vendored icon assets (Phosphor fill + Lucide) + manifest + LICENSE *
  aesthetic.py  # aesthetic audit: whitespace/font-size hierarchy/color count/contrast/consistency → scored report *
  autopick.py   # automatic pick: content structure → recommended theme + per-slide layout + reasoning *
  feedback/     # feedback learning (v3 numpy MLP): registry/vector/preprocess/train/validate/infer/status/app *
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
[`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md); each icon set's `LICENSE-*.txt` is in
`src/offipy/assets/icons/`.

## Feedback and issues

- **Bugs / feature requests**: file them at [GitHub Issues](https://github.com/Zn070515/offipy/issues),
  attaching `offipy check --json` output and a minimal repro.
- **Pre-release versions**: smoke-test TestPyPI with `scripts/pypi_smoke.py --index https://test.pypi.org --version <pre-release version>`
  (downloads the wheel via the TestPyPI JSON API with a dual sha256 check; see "Pre-release numbering policy" in the CHANGELOG).
  Pre-release versions are for validation only; before the initial stable release,
  the top three — `__version__` / tag / CHANGELOG — must stay consistent.
