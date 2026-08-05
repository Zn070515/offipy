> [中文](audit-baseline.md)

# PPTX Baseline Regression (compare_pptx)

Pairs with [audit](audit.en.md) to answer "**did the new change introduce new problems**":
compare a **baseline** PPTX against a **candidate** PPTX, aggregating added / resolved / changed
findings plus shape add / remove / move / resize / text changes.

- **Baseline**: the previously approved artifact (or a historical version).
- **Candidate**: the artifact after this change.
- Only problems the candidate **adds or worsens relative to the baseline** trigger the gate;
  pre-existing baseline issues **pass**.

## CLI

```bash
# regression: exit code 1 when added/worsened reaches MID
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID

# regression + HTML report
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID \
  --format html --out regression.html
```

- `--baseline PATH`: entering compare mode when given (default `--format text`).
- `--fail-on-new HIGH|MID|LOW`: **only for candidate added/worsened issues**. Reaching that
  severity → exit code 1.
- `--baseline` requires `--fail-on-new` (`--baseline` alone without `--fail-on-new` → exit code 2);
  `--fail-on-new` without `--baseline` is also misuse → exit code 2.
- Compare-mode exit codes match the normal audit: 0=threshold not reached / 1=gate hit /
  2=argument or input error / 3=dependency or parse error.

## Python API

```python
from offipy import compare_pptx, AuditConfig

diff = compare_pptx(
    "baseline.pptx",
    "candidate.pptx",
    audit_config=AuditConfig(safe_margin_in=0.2),
)

print(diff.gate_severity())   # Severity.MID / LOW / HIGH / None (no added/worsened)
for f in diff.added_findings:
    print("added", f.rule_id, f.severity.name)
for f in diff.resolved_findings:
    print("resolved", f.rule_id)
print(diff.to_json())         # fully JSON-safe
```

## Shape matching chain

Per page, in order, **the first matching level claims the pair** (`_match_slide` in `compare.py`):

1. same `shape_id` on the same page;
2. `name + shape_type`;
3. normalized-text hash + center distance (≤1.0 inch, nearest);
4. image content sha256 (only available for PICTURE).

Anything still unmatched lands in `unmatched_baseline` / `unmatched_candidate`
(low confidence; the caller decides whether to warn).

> **shape_id must not be assumed permanently stable across edits** — it is only the first-priority
> match key; inserting/deleting shifts later shape_ids, so the later chain levels must back it up.

## PptxDiffReport fields

| field | meaning |
|-------|---------|
| `baseline_path` / `candidate_path` | both paths |
| `baseline_sha256` / `candidate_sha256` | both file fingerprints |
| `baseline_slide_count` / `candidate_slide_count` | slide counts (`added_slides` / `removed_slides` are the delta properties) |
| `baseline_findings` / `candidate_findings` | each side's full audit results |
| `added_findings` | findings added by the candidate |
| `resolved_findings` | findings present in the baseline, gone in the candidate |
| `changed_findings` | same finding on a matched shape whose severity changed (`ChangedFinding.worsened` marks an increase) |
| `added_shapes` / `removed_shapes` | shape add/remove |
| `moved_shapes` / `resized_shapes` / `text_changes` | position / size / text changes (with old & new geometry, inches) |
| `unmatched_baseline` / `unmatched_candidate` | objects the matching chain could not claim (low confidence) |
| `warnings` | union of both sides' parse warnings |

**Gate semantics**:
- `new_or_worsened` → the list of candidate added-or-worsened findings (`--fail-on-new` only looks at these).
- `gate_severity()` → the highest severity among candidate added-or-worsened findings;
  `None` if none (gate does not trigger).

```python
# manual check: did the candidate introduce MID+ added/worsened issues?
if diff.gate_severity() is not None and diff.gate_severity() >= Severity.MID:
    raise SystemExit("candidate introduced MID+ added/worsened issues")
```

## Details

- Severity comparisons **use integer values** (`Severity` is an `IntEnum`); string comparison is forbidden.
- In `changed_findings`, only `worsened=True` counts as a regression and enters
  `new_or_worsened` / `gate_severity`; a **downgrade** change does not trigger the gate (recorded only).
- A finding's match key is `(rule_id, primary(slide,shape_id), secondary)`; if either the primary
  or secondary shape fails to match, it is treated as added.

## Fixed verification corpus (tests/fixtures/audit/)

- `baseline.pptx` + `candidate.pptx` produced by the generator: the candidate adds an
  out-of-bounds `bounds.partial` MID, resolves an edge-adjacent `margin.right`, and includes
  move / resize / text changes plus shape add/remove.
- `tests/test_audit_fixtures.py::test_compare_finds_added_and_resolved` asserts:
  `added_findings` contains `bounds.partial` MID; `resolved_findings` contains `margin.right`;
  `added_shapes==2` / `moved_shapes==2` / `resized_shapes==1` / `text_changes==1`;
  `gate_severity()==MID`.

## CI usage

```bash
# only block candidate added/worsened MID+ issues (pre-existing baseline issues pass)
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID

# stricter: block any added/worsened
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new LOW

# machine-readable: hand the diff to downstream
offipy audit candidate.pptx --baseline baseline.pptx --fail-on-new MID --format json
```

> `rule_id` is the stable machine key; `message` is for humans only. In `--format json`,
> `max_new_severity` is the serialized `gate_severity()` (`"LOW" / "MID" / "HIGH" / null`).
