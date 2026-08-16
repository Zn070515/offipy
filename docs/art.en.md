# Art Analysis (offipy.art)

`offipy.art` is a **pure-stdlib, deterministic, advisory-only** visual / typographic quality analysis: no AI, no
Microsoft Office, and `import offipy` does not load python-pptx. It abstracts a slide as a "scene (ArtScene)"
and evaluates 5 dimensions (hierarchy / composition / typography / color / media) with deterministic rules,
producing a per-slide report.

- **Advisory only, never blocks**: every finding carries `confidence` and `severity`, but the art layer has
  **no total-score gate** and makes no pass / fail judgment — that trade-off is left to the caller (e.g. the
  deck pipeline consumes it as a post-generation quality reference).
- **Deterministic**: the same input always yields the same output — no randomness, no models, no network.
- **Evidence honesty**: dimensions with insufficient evidence degrade to `insufficient_evidence` rather than
  fabricating numbers and false-reporting.

## Install & dependencies

The art layer has **zero extra dependencies** (pure stdlib); `import offipy` is enough:

```python
import offipy
offipy.analyze_scene   # its presence proves the art layer is ready
```

Two evidence sources (below) are opt-in as needed:

- `measurements` (the HTML→PPTX pipeline's DOM measurement JSON) — shipped with the `deck` pipeline;
- `pptx` (a PPTX geometry audit report) — parsing `.pptx` requires python-pptx (`pip install "offipy[deck]"`),
  but it is only loaded when a file is actually parsed.

## Quick start

```python
from offipy import build_scene, analyze_scene, render_markdown

# 1) Build the scene: measurements (real browser pixel evidence) primary, geometry audit secondary
scene = build_scene(
    measurements="out/report_audit/_cache/measurements.json",  # where the deck pipeline writes it
    pptx="out/report.pptx",
)
# 2) Analyze
report = analyze_scene(scene, profile="balanced")
# 3) Report
print(render_markdown(report))
```

Combined entry point (geometry audit + art analysis in one call):

```python
from offipy import analyze_deck

report = analyze_deck(
    pptx="out/report.pptx",
    measurements="out/report_audit/_cache/measurements.json",
    profile="consulting",
)
print(report.art.slides[0].by_dimension("color").status)   # assessed / insufficient_evidence
for s in report.art.slides:
    for d in s.dimensions:
        for f in d.findings:
            print(s.slide_index, d.dimension, f.rule_id, f.confidence)
```

With only `pptx=` (no measurements), `_ShapeRecord` enrichment provides font-size / font-family /
foreground / background / opacity / fill_kind evidence (#128); but when most runs inherit the theme
font (no explicit size) coverage may stay low, so the hierarchy / typography / color dimensions can
still be `insufficient_evidence` (+ `art.evidence.limited` warning), while pure-geometry rules
still run:

```python
report = analyze_deck(pptx="external.pptx", profile="balanced")
```

## Evidence sources & coordinate conventions

An ArtScene is built from two evidence sources, merged as "measurement-first, audit-secondary":

| Source | Evidence | Unit | Notes |
|--------|----------|------|-------|
| `measurements` (MeasurementAdapter) | color, font size, natural size, text, font family | px (normalized to a [0,1] score) | real browser-pixel evidence from the HTML→PPTX pipeline; role vocabulary title/body/subtitle/image/shape |
| `pptx` (PptxAuditAdapter) | geometry (position / size / role classification), text, font size / font family / foreground / background / opacity / fill_kind | pt (normalized to a [0,1] score) | `audit_pptx` geometry snapshot + `_ShapeRecord` enrichment (#128); when most runs inherit the theme font (no explicit size) font/color evidence is limited; role vocabulary background/header/footer/page_number/title/content/decoration/unknown |

Merge (`merge_scenes`) pairs elements one-to-one by **text as strong corroboration + geometry as
fallback**: same normalized role and same text → identity 0.8; cross-vocabulary same text → 0.7; both have
text but differ → not merged; at least one side has no text → geometric proximity (center distance ≤ 0.2) +
role bonus. Unmatched audit elements are kept with an `art.merge.unmatched` warning — never silently dropped.

Coordinates are normalized to slide-width / slide-height fractions ([0,1]); `font_size_norm = font_size / page
height` (dimensionless). Rules compare only normalized values and never raw sizes across sources.

## grade / confidence / evidence_coverage are kept separate

The semantics of each dimension (DimensionAssessment) and each finding are strictly separated; none
impersonates another:

- **grade**: `excellent / good / attention / poor` — quality only. Accumulated from the severity-weighted
  scores of the dimension's assessed rules (LOW=0.5 / MID=1.5 / HIGH=3.0), bucketed at `(0, 1.0, 2.5)`.
- **confidence**: `(0, 1]` — credibility only. Findings with better evidence coverage and closer to a
  threshold edge are more credible; experimental rules are forced to `conf ≤ 0.4` (above the 0.35 grade floor, so they participate in grade at low weight).
- **evidence_coverage**: `[0, 1]` — evidence coverage only. Reported by each rule itself as eligible/covered,
  never a guessed constant; **coverage gating is decided per rule**: only rules with eligible>0 and
  covered/eligible ≥ 0.5 keep their findings; the dimension is `insufficient_evidence` only when no rule
  qualifies; gated rules emit an `art.rule.insufficient_coverage` warning.

Dimensions also carry an optional **`reliability`**. In addition to an explicit `ev.reliability`, a
dimension's reliability is derived from the min of its rules' finding `evidence_reliability`
(experimental still excluded).

`experimental_score` (0-100) is only returned when ≥ 3 dimensions are assessed; it exists purely for
ranking / comparison and is **never claimed as an objective aesthetic score**. When a valid feedback
model exists (see [Feedback learning system](usage.en.md)), the score source changes from the formula
to the learned `quality.score` — the semantics are unchanged: still only for ranking / comparison and
never claimed as an objective aesthetic score; `include_experimental_score` stays opt-in.
`quality.score` is contributed only by findings that pass the **evidence gate** (the rule has ≥ 3
valid labels, see #111) and are **confident (not abstained) / in-distribution (not OOD)** (#122);
the value is mapped from the ensemble-mean worth via calibration (worth_scale normalization) (#122).
Without a model, or when every finding abstains / is OOD, the current formula is kept.

## Dimensions & rules

| Dimension | Rules | Notes |
|-----------|-------|-------|
| hierarchy | no_focus / focus_conflict / active_title | whether focus is clear and whether it conflicts with the title |
| composition | off_balance / corner_cluster / spacing_drift | center of gravity, corner clustering, spacing drift |
| typography | tiny_text / many_families / flat_scale | font too small, too many families, flat scale |
| color | no_accent / accent_flood / low_contrast | missing accent, accent overload, foreground/background contrast |
| media | distorted_image / oversized_image / image_overlap | distorted, oversized, mutually overlapping images |

All rules are **deterministic**: the same scene always yields the same result. All `rule_id`s are frozen
(including 6 experimental ones, `conf ≤ 0.4`, participating in grade at low weight).

- The `tiny_text` / `title_too_small` / `no_focus` findings carry learnable details
  (`font_size_norm` / `ratio_vs_min` / `focus_ratio`) (#145).
- Accent-color assessment uses run-level colors (weighted by text length) with scope unified with the
  palette (`_SKIP_ROLES` + zero-area exclusion) (#156).

## Built-in profiles

`offipy.profile_names()` → `['balanced', 'consulting', 'academic', 'technology', 'event']`. A profile
controls the enabled rule set, experimental rules and thresholds; `get_profile(name)` is readable and
extendable.

## Deck consistency

Beyond per-slide rules, `analyze_scene` also evaluates cross-slide consistency (`assess_deck`): grouped by
role (≥ 3 elements) it checks spacing / font-size / background-color drift and produces
`report.deck_findings`.

## Integration with the deck pipeline

`deck.render_with_quality_report(html, out=..., audit_mode=..., fail_on=..., profile=...)`:
after HTML→PPTX generation it also produces the geometry audit (`audit_pptx`) and the art analysis
(`build_scene` + `analyze_scene`) and returns a `QualityRenderResult` (with `art_report` / `deck_quality`) —
quality-on-generate. `audit_mode="report"` only reports; `"strict"` exits non-zero only when the `fail_on`
threshold is reached.

## Known boundaries

- **PptxAuditAdapter enrichment boundary**: with only `pptx=` given, `_ShapeRecord` enrichment provides
  font-size / font-family / foreground / background / opacity / fill_kind evidence (#128); but when most
  runs inherit the theme font (no explicit size) coverage may stay low, so the hierarchy / typography /
  color dimensions can still be `insufficient_evidence`, not a false finding; pass `measurements=` for
  full pixel evidence.
- **RenderedSlide (PNG / slides_dir pixel-level) analysis deferred to v0.12.1**: `build_scene(slides_dir=...)`
  explicitly rejects it (`InvalidArgumentError`).
- **Unmatched elements in the dual-source merge**: audit decoration elements not modeled by measurements
  (e.g. decorative lines) are kept with a warning — not silently dropped, and not diluting the evidence of
  already-matched elements.
