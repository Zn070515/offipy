> [中文](feedback.md)

# Feedback API

### `train`

Train the feedback learning system offline: read ~/.offipy/art_feedback.jsonl → encode against the current FEATURES schema → build pairs (same rule×profile: fixed > accepted) → train a numpy MLP → atomically write art_feedback_model.json. With too few samples or no valid samples it returns a status instead of failing (does not delete an existing model). Requires numpy: pip install "offipy[feedback]".

- **Parameters**: `feedback_dir: str`, `seed: int`
- **Returns**: `dict`
- **Flags**: normal operation

---

### `status`

Feedback learning status: sample count, pairing potential, current model state (none/valid/expired/stale/corrupt). Read-only over local data — no training, no writes.

- **Parameters**: `feedback_dir: str`
- **Returns**: `dict`
- **Flags**: read-only

---

### `append`

Append one feedback label: how a user disposed of a finding for a rule (fixed = should fix, accepted = rule is right, ignored = irrelevant). Written to the feedback_dir JSONL (default ~/.offipy when feedback_dir is omitted), consumed by feedback train. features is a flat feature snapshot (encode_features output); the CLI accepts a JSON string.

- **Parameters**: `profile: str`, `rule_id: str`, `action: str`, `severity: str`, `slide_index: int`, `message: str`, `source: str`, `feedback_dir: str`, `ts: str`, `features: any`, `feature_schema_version: str`
- **Returns**: `dict`
- **Flags**: normal operation

---

### `recommend`

Read-only recommendations: run art analysis + learned inference on a .pptx and return adjusted findings and deterministic suggestions (no document writes, no feedback-store writes). Requires a valid model: without one / expired / corrupt it raises explicitly (no silent v2 fallback). --json is accepted (generic dispatch always outputs JSON).

- **Parameters**: `pptx: str`, `feedback_dir: str`, `profile: str`, `json: bool`
- **Returns**: `dict`
- **Flags**: read-only

---

### `apply`

Persist learned rule.delta to the profile store (default ~/.offipy/art_profiles.json), so `deck audit --profile <name>` (without --feedback-dir) also reflects learned adjustments. Requires a valid model.

- **Parameters**: `profile: str`, `feedback_dir: str`
- **Returns**: `dict`
- **Flags**: normal operation
