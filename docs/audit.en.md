> [中文](audit.md)

# PPTX Quality Audit

`offipy audit` is a set of **static geometry quality gates**: without opening PowerPoint and
without depending on Microsoft Office, it directly parses `.pptx` (ZIP+XML) structured shape
extraction, checks out-of-bounds / edge-adjacent / overlapping / text overflow / autofit risks,
produces text / json / markdown / html reports, and blocks non-conforming artifacts by severity
threshold.

- Emits a **stable `rule_id`** (not a natural-language message) — users / CI depend on it for automation.
- Configurable suppressions, readable suppressed (why it wasn't reported), readable warnings (what couldn't be parsed).
- Pairs with `compare_pptx` ([baseline regression](audit-baseline.en.md)) for "did the new change introduce new problems" regression.
- Pairs with `deck render_with_report` for "HTML→PPTX generate-and-gate".

## Installation & dependencies

The audit core **does not depend on Microsoft Office** (pure parsing, no COM). Reading `.pptx`
needs python-pptx:

```bash
pip install "offipy[deck]"    # includes python-pptx
```

`import offipy` / `import offipy.audit` do **not** load python-pptx (hard lazy-import constraint);
it is only needed when actually parsing a file.

## Quick start

```bash
# text report (default)
offipy audit deck.pptx

# report + gate: exit code 1 when reaching HIGH (for CI)
offipy audit deck.pptx --fail-on HIGH

# JSON / Markdown / single-file HTML (SVG canvas, filterable)
offipy audit deck.pptx --format markdown
offipy audit deck.pptx --format html --out audit.html --slides-dir export/
```

```python
from offipy import audit_pptx

report = audit_pptx("deck.pptx")
print(report.max_severity)      # Severity.HIGH / MID / LOW / None
print(report.to_markdown())
for f in report.findings:
    print(f.rule_id, f.severity.name, f.message)
```

## CLI reference

```
offipy audit <file.pptx>
  --format text|json|markdown|html    default text (html writes <stem>.audit.html by default)
  --out PATH                          report output file; text/json/markdown default to stdout
  --fail-on HIGH|MID|LOW              audit reaching that severity → exit 1 (default HIGH)
  --baseline PATH                     when given, run regression comparison (see audit-baseline.en.md)
  --fail-on-new HIGH|MID|LOW          compare mode: candidate added/worsened reaching that severity → exit 1
  --safe-margin FLOAT                 safe margin (inches, default 0.2)
  --bounds-tolerance FLOAT            out-of-bounds tolerance (inches, default 0.01)
  --no-full-bleed-ignore              disable full-page background suppression
  --no-repeated-decoration-ignore     disable repeated-decoration suppression
  --no-page-number-ignore             disable page-number suppression
  --no-header-footer-ignore           disable header/footer suppression
  --slides-dir PATH                   html only: PNG slide background directory (slide-<n>.png)
  --show-suppressed                   suppressed items are listed by default; flag kept for compatibility
  --debug                             print full traceback on failure
```

**Exit codes**:

| code | meaning |
|------|---------|
| 0 | threshold not reached (audit passed) |
| 1 | succeeded but reached `--fail-on` / `--fail-on-new` (gate hit) |
| 2 | argument or input error (missing file, misused `--fail-on-new`, etc.) |
| 3 | dependency or parse error (missing python-pptx, corrupted ZIP/XML) |

`_audit_main` catches all expected exceptions and converts them to exit codes, never colliding
with the other commands' `OffipyError → 1`.

## Python API

```python
from offipy import (
    Severity, AuditConfig, AuditFinding,
    PptxAuditReport, PptxDiffReport, audit_pptx, compare_pptx,
)

report = audit_pptx(
    "deck.pptx",
    AuditConfig(
        safe_margin_in=0.2,          # safe margin (inches)
        bounds_tolerance_in=0.01,     # out-of-bounds tolerance (inches)
        ignored_shapes={(1, 42)},     # explicit user suppression: slide 1 shape #42
        ignored_regions=[(0.0, 6.5, 10.0, 1.0)],  # suppress a whole bottom region (inches x,y,w,h)
    ),
)
```

### Severity

`LOW=1 / MID=2 / HIGH=3` (`IntEnum`). **Comparisons must use integer values**
(`f.severity >= Severity.HIGH`); string comparison is forbidden. Serialized output is
`"LOW"/"MID"/"HIGH"`.

### Stable rule_id

| rule_id | severity | meaning |
|---------|----------|---------|
| `geometry.bounds.partial` | MID / HIGH | shape partially out of slide bounds |
| `geometry.bounds.off_canvas` | LOW / MID | fully off-canvas (staging / animation / design residue) |
| `geometry.margin.left/right/top/bottom` | LOW | content near an edge, spacing < safe margin |
| `geometry.overlap.partial` | LOW / MID | shapes partially overlap (coverage ratio > 0.5, by smaller shape) |
| `geometry.overlap.covered_text` | MID / HIGH | one shape fully covers another; text covered by picture/chart is HIGH |
| `text.fit.horizontal` | LOW / MID | text exceeds the text box horizontally (nowrap single line too wide, etc.) |
| `text.fit.vertical` | LOW / MID | text exceeds the text box vertically (explicit multi-line too tall / no available space) |
| `text.autofit.shrink` | MID / HIGH | normAutofit shrinks the font, possibly below the 8pt minimum readable |
| `text.autofit.grow` | MID / HIGH | spAutoFit grows the shape, possibly out of bounds / colliding |

### Report models

- `PptxAuditReport`: `max_severity` (`None` with no findings), `findings` / `suppressed` /
  `warnings` / `shapes` (per-shape geometry snapshot), `to_dict()` / `to_json()` /
  `to_markdown()` / `to_html(slides_dir=...)`. JSON output is **fully safe** (no Enum / Path / set).
- `AuditFinding`: `rule_id` / `kind` / `severity` / `message` / `primary` / `secondary`
  (two shape refs) / `details` / `confidence`. `message` is Chinese natural language;
  `rule_id` is the stable key.
- Coordinate unit: **slide-absolute inches** (group children already absolutized).

## Rules

A registry drives the rules (`DEFAULT_RULES`), executed in order:
Bounds → Margin → Overlap → TextFit → Autofit.

### Bounds

- Any edge beyond the page and past `bounds_tolerance_in` → `geometry.bounds.partial`.
- Large overshoot (`out_ratio > 0.5` or max overshoot > 0.25× the page long edge) → **HIGH**, else **MID**.
- No intersection with the canvas → `geometry.bounds.off_canvas` (area ≥1 in² → MID, else LOW).
- A shape already flagged for bounds on an edge is **not** flagged for the same-direction margin (no double report).
- Hidden / group / geometrically unresolvable objects are skipped.

### Margin

- `safe_margin_in=0.2`; ordinary content spacing < 0.2 → `geometry.margin.*` LOW.
- **Suppressions** (into suppressed with a reason): full-page background (`full_bleed`),
  page number (`page_number`), header/footer (`header_footer`), repeated decoration
  (`repeated_decoration`), user ignore (`user_shape` / `user_region`).
- Connectors / hidden / group are skipped.

### Overlap

- Per-page bbox quick-reject + O(n²); `ratio = overlap_area / min(shape areas)`, reported only when > 0.5.
- Skipped: connectors / hidden / group / full-page background / tiny decoration points
  (area < 0.0025 in²) / parent-child·ancestor pairs / geometrically unresolvable.
- **Pair classification**: text inside a filled AutoShape → card container, suppressed as
  `intentional_containment`; text fully covered by a picture/chart → HIGH; same Group with a
  reasonable z-order → downgraded severity.
- Rotated shapes use AABB approximation (confidence 0.5, message notes "rotated bounding-box approximation").

### TextFit

- Available area is **reduced by the TextFrame's margins first**.
- Reported when: zero width/height, a nowrap single line clearly too wide, explicit
  line-count × line-height clearly exceeding available height.
- Font metrics: **Pillow first** (font file locatable → confidence 0.8); **when no font file is
  found, fall back to character weights** (CJK=1.0 / ASCII=0.5 / space=0.35 → confidence 0.4,
  message notes "low-confidence character estimation").
- Page-number / header / footer small text is skipped (naturally compact).

### Autofit

The two modes are **kept separate** (not uniformly downgraded):
- `normAutofit` (shrink font to fit shape) → `text.autofit.shrink`: records original font size /
  fontScale / estimated size; estimated size < 8pt → HIGH.
- `spAutoFit` (grow shape to fit text) → `text.autofit.grow`: HIGH if the grown shape may overflow,
  else MID.

## Roles & suppressions (suppressed)

`suppressed` is a **suppression record with a reason**, not silently dropped. Common reasons:

| reason | trigger |
|--------|---------|
| `full_bleed` | full-page background (≥90% coverage + near-center + low z-order + no text) |
| `page_number` | pure number + bottom 15% region + small size (or slide-number placeholder) |
| `header_footer` | header/footer placeholder; or repeated >60% across pages + top/bottom region |
| `repeated_decoration` | decoration whose fingerprint repeats >60% of pages |
| `intentional_containment` | text inside a filled AutoShape (card container) |
| `user_shape` / `user_region` | user explicitly suppressed via `ignored_shapes` / `ignored_regions` |

**Not all pure-number short text is globally ignored** — only small pure numbers in the bottom
region count as page numbers.

## confidence semantics

| confidence | meaning |
|------------|---------|
| 1.0 | exact geometry, no heuristics |
| 0.8 | Pillow font metrics used for text-width estimation |
| 0.5 | AABB approximation of a rotated shape (message notes it) |
| 0.4 | character-weight fallback (message notes "low-confidence character estimation") |

## warnings (parse exceptions)

`warnings` record items the parse layer cannot handle precisely: `group.no_transform`
(group missing `a:xfrm` → children cannot be precisely positioned → rules requiring exact
position are skipped). These cases are **not** silently treated as zero rotation.
This is the only warning source in v0.11.

## Known limitations & false-positive control

- **v0.11 does not run text-fit checks on** table cells / SmartArt / text inside charts /
  WordArt / vertical text / complex bullets and custom line spacing — they are silently
  skipped (no hard-fail, no false positive). Structured unsupported warnings
  (`textfit.table_unsupported` etc.) are planned for 0.11.1.
- Rotated groups / flips use AABB approximation for bounds/margin/overlap (the geometric shape
  occupies the same place; flipping only affects content orientation).
- No promise of zero false positives on every PPT — the fixed verification corpus
  (`tests/fixtures/audit/`) guarantees connector / hidden / rotate / flip / group are not
  misjudged, full-page background / page numbers do not false-report margin, and reasonable card
  containment does not false-report overlap; the rest is tuned with `ignored_shapes` /
  `ignored_regions` / `--no-*-ignore`.

## Deck generation gate (render_with_report)

After HTML→PPTX rendering, audit immediately and decide release by mode:

```python
from offipy import deck

result = deck.render_with_report(
    "deck.html", audit_mode="strict", fail_on=deck.Severity.HIGH,
)
# passed: replaces the target and returns a RenderResult (output_path + audit_report)
# failed: raises deck.AuditGateError (report on the exception, temp file cleaned, old target untouched)
```

```bash
offipy deck make --html deck.html --out deck.pptx --no-open \
  --audit-mode strict --fail-on HIGH --audit-report deck.audit.json
```

- `report` (default): render → audit → replace → return `RenderResult`.
- `strict`: render → audit → max severity ≥ `fail_on` → raise `AuditGateError`
  (report is written to disk first, the old `.pptx` is not corrupted); otherwise → replace.
- Atomic replace: conversion writes a same-directory temp file first; only replace with
  `os.replace` after the audit passes.

## CI usage

```bash
# block: any HIGH issue
offipy audit deck.pptx --fail-on HIGH

# regression: only block candidate added/worsened MID+ issues (pre-existing baseline issues pass)
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID
```

Hand results to downstream with `--format json`; `rule_id` is the stable machine key,
`message` is for humans only.
