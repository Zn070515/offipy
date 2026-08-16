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

# Feedback learning CLI (train/status/append/recommend/apply, dedicated subparser)
offipy feedback --help                 # 5-op guide for feedback learning
offipy feedback train [--feedback-dir <dir>]        # offline model training
offipy feedback status [--feedback-dir <dir>]       # sample/model status
offipy feedback append --profile <p> --rule-id <id> --action fixed|accepted|ignored --severity LOW|MID|HIGH [--features '<json>'] [--feedback-dir <dir>]
offipy feedback recommend --pptx <deck.pptx> --feedback-dir <dir> [--profile <p>] [--json]
offipy feedback apply --profile <p> [--feedback-dir <dir>]
```

Destructive ops need a target: `--doc_id <doc_id>` / `--follow-active` / `--expected-target '<json>'`.
Boolean parameters use `--key true/false`: `--overwrite true`. Structured values can be
passed with `--payload '{"...": ...}'`. Parameter names use underscores (e.g.
`--range_addr`, `--doc_id`); types are declared in `schema.py` and converted automatically.
The `feedback` subcommands accept dual spellings: `--feedback_dir`/`--feedback-dir` and `--rule_id`/`--rule-id` are equivalent.
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

A `<div class="drawio" data-drawio="arch.drawio">` inside a `data-pptx-slide`
`<section>` is injected as editable PowerPoint shapes after render (needs the visual
audit measurements). Relative `data-drawio` paths are rewritten to `file://` absolute
URIs when the HTML is staged to a temp dir, so they keep resolving; node `fontSize`
is scaled with the container (default 12pt), keeping the text hierarchy.

Multipage `.drawio` files must name the page in the deck: use
`data-drawio-page="N"` (1-based) or a page name (case-insensitive), e.g.
`<div class="drawio" data-drawio="arch.drawio" data-drawio-page="2">`. Omitting it
for a multipage source raises an error instead of silently taking the first page.
Orthogonal/curved edges render as polylines along their waypoints (arrows kept),
and `strokeWidth`, `rotation`, and `dashPattern` (space-separated pairs) propagate
to the generated shapes.

### Media fidelity

Media fidelity (#141): `<audio>` cannot be expressed in PPTX, so it is silently dropped
during conversion with a warning; `<video>` is embedded as a static first-frame image (no
native playback), also with a warning; an uncached webfont skips subset embedding and is
replaced by PowerPoint when opened — these warnings surface through the deck quality
report (`deck.media.audio_dropped` / `deck.media.video_static` / `deck.font.substituted`).
`<a href>` links are written as real hyperlinks (clickable).

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

### Art / audit evidence layer

- **distorted_image is now judged by decoded vs rendered size** (#126): the rule's
  `natural_ratio` / `physical_ratio` are now based on the image **decoded size** (img uses
  `naturalWidth/Height`; SVG scales from its viewBox) versus the **rendered size** (CSS layout
  width/height), so real stretch drift is detected; the old implementation took both ratios
  from the rendered size, so drift was always ≈0 and the rule was effectively dead. This
  semantic change invalidates historical FEATURES samples and moves `feature_schema_version()`
  from 2 to 3.
- **PPTX-only enrichment** (#128): when only `pptx=` is passed, `audit_pptx` / `analyze_scene`
  now parse font size / font family / foreground / background / opacity / fill_kind from
  `_ShapeRecord`, so the hierarchy / typography / color dimensions gain font and color evidence
  coverage instead of starting from 0; schemeClr / sysClr (theme-color references) are still
  not resolved, and pure pixel evidence (PNG / measurements) still needs an additional source.
- **Element opacity** (#137): the measurement's element-level opacity flows into
  `ArtElement.opacity` (0-1, None = no evidence); the PPTX-only path merges run foreground
  alpha and shape fill alpha with min (the most transparent part decides visibility).
- **fill_kind marking** (#140): measurements record `fill_kind` (`solid` / `gradient` /
  `shadow` / `image`), so rasterized decoration such as gradients / shadows is recognized by
  the audit instead of being mistaken for a plain color block.

## Feedback learning system (v0.18)

Three layers of feedback semantics (pinned here to avoid confusion):
- `offipy.feedback` (v1): dimension weights, `dimension_weights()`, `~/.offipy/feedback.jsonl`
- `offipy.art.feedback` (v2): per-rule ±1, `recommend_adjustments` → `feedback_severity_adjustments`
- `offipy feedback` (v3, this system): learnable numpy MLP, `feedback_train` /
  `feedback_status` / `feedback_append` / `feedback_recommend` / `feedback_apply`

Training: `offipy feedback train` (reads `~/.offipy/art_feedback.jsonl` → trains →
writes `~/.offipy/art_feedback_model.json`). With insufficient / no valid samples it returns a
status JSON instead of erroring and does not delete an existing model (F2-E). **Numerical gates
(#112)**: non-finite loss → `training_diverged`; constant output (output_std < 1e-6) →
`model_collapsed`; a bad model is never written and the old one is kept atomically; training
uses global gradient clipping. Requires numpy: `pip install "offipy[feedback]"`.

**Learning quality (#115-#122)**:
- **Standardized preprocessing (#118/#120)**: zero-variance feature drop + high-correlation
  (|r|≥0.99) dedup + global z-score fitted on the training set; mean/scale/kept persist to the
  model.json `preprocessing` block (model schema v2), and inference uses the same transform
- **Ensemble K=5 (#122)**: multi-seed members averaged to reduce variance + worth calibration
  (quality_score normalization); **abstain** (findings with near-zero |worth| / high member
  disagreement are not shifted) + **OOD** (features out-of-distribution are not shifted),
  conservative fallback to v2
- **Sample-level repeated stratified CV (#119)**: per-rule stratified repeated 5-fold, never a
  pair-level split; a 95% confidence lower bound at chance → `poor_generalization` soft flag
  (recorded only, never rejects)
- **Capacity-adaptive warning (#121/#134)**: capacity adapts to the independent sample count
  (H≈√n); `samples_per_param` is graded ok/warn/critical (recalibrated in #134 to spp≥1 ok /
  ≥0.25 warn / <0.25 critical), plus `suggest_n` (estimated samples still needed to reach ok;
  recorded only, never rejects)
- **Per-rule sample diagnosis (#117/#152)**: on insufficient_pairs, returns per-rule
  `{fixed, accepted, pairs, single_direction, suggest}` actionable suggestions; `status` and the
  `train` success return also carry `per_rule`
- **Model output saturation detection (#151)**: P5-P95 span of per-sample quality scores (fixed
  reference scale) too small → `saturation` soft flag (recorded only, never rejects); the analyze
  learning pass emits a `feedback.model.saturated` warning
- tiny_image feature completion + FEATURES schema v1→v2 bump (#115)

Append labels: `offipy feedback append --profile <p> --rule_id <r> --action fixed
--severity MID --features '<json>' --feedback_dir <dir>` (writes to that directory's JSONL for
`train` to learn; severity is limited to LOW/MID/HIGH).

Status: `offipy feedback status` (sample count / valid sample count / pairing potential /
model none|valid|stale|corrupt|expired; when the model is valid it additionally surfaces
`effective_dims` / `samples_per_param` / `poor_generalization` / `saturation`; every branch
carries `excluded` (filtered-record breakdown) and `per_rule` (per-rule sample diagnosis).
`stale` = schema matches but kept indices are out-of-bounds/missing/non-numeric (a bump forgot
to retrain, #150); `corrupt` = schema matches but the weights fail to rebuild (#147); neither
masquerades as a valid model).

Consumption rule (#113): learning **consumption** requires an explicit `feedback_dir` —
`analyze_scene(feedback=True)` without a dir → `InvalidArgumentError`; `learned_adjustments`
without a dir → returns None, and never silently loads the global `~/.offipy` model. The write
side (train/append) still defaults to `~/.offipy`.

CLI consumption channel (#114/#116):
- `offipy deck audit --feedback-dir <dir>`: applies feedback learning during audit (loads
  records/model from the given dir); **when `--feedback-dir` is given the report / `--json`
  output passes through `experimental_score`** (#116)
- `offipy deck make --export-png <dir>`: export PNG feedback dir (the old name `--feedback` is
  a deprecated alias, semantics unchanged — it still only exports the dir, it does not trigger
  learning consumption)

Inference consumption points:
- `rule.delta.<rule_id>`: mean worth of historical records → ±1 adjustment →
  `feedback_severity_adjustments`; also lands in `report.feedback_adjustments`, visible in the
  deck audit JSON / text "模型调整 (rule.delta)" segment (#159)
- `finding.severity_shift`: a post-processing pass in analyze, applied only to rule-computed
  (no-override) findings; **per-rule evidence gate (#111)** — only rules with ≥ 3 valid labels
  get shifted; 0-label rules are not cross-rule-generalized; findings the model is uncertain
  about (abstain) or that are out-of-distribution (OOD) are not shifted (conservative fallback
  to v2); when a shift applies the finding is marked override + traced in `details.feedback`
  (before/after/worth/shift, #157), and the dimension grade is re-derived from post-shift
  findings (#132)
- `quality.score`: replaces `experimental_score` (only computed when
  `include_experimental_score=True`); **requires ≥ 3 assessed-dimension worths** (below that no
  score is written, #158); same evidence gate — contributed only by gated (shiftable) findings;
  the value is mapped from the ensemble-mean worth via calibration (worth_scale normalization),
  and the source is annotated in `experimental_score_mode` (`worth_sigmoid`; grade-mean source
  is `grade_mean`, #130); coverage of the confident subset is reported via
  `quality_score_coverage` (covered/total/abstain/ood, #133); with no valid model a
  `feedback.model.unavailable` warning is emitted instead of a silent v2 fallback (#158)

Cold start: no model / expired model (feature_schema_version / model schema mismatch) / stale
model (schema matches but kept indices out-of-bounds, #150) / corrupt (weights fail to rebuild,
#147) / numpy not installed → full fallback to v2 behavior (`recommend_adjustments`). Deleting
model.json returns to v2. The learning system is a detachable enhancement and never relaxes
the audit hard gate.

recommend / apply (#160):
- `offipy feedback recommend --pptx <deck.pptx> --feedback-dir <dir>`: read-only recommendations —
  runs analysis + learned inference, projects `adjusted_findings` / `suggestions`, writes nothing
  to documents or the feedback store; with no valid model it raises explicitly (no silent v2
  fallback).
- `offipy feedback apply --profile <p>`: persists the learned rule.delta to
  `~/.offipy/art_profiles.json` — afterwards `deck audit --profile <name>` (without
  `--feedback-dir`) also reflects the adjustment (the default store only takes effect for a
  profile once adjustments have been applied to it).

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
