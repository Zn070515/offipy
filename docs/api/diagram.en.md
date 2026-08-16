> [中文](diagram.md)

# Diagram API

### `build`

Convert a Mermaid/drawio source file into an editable PPTX (16:9 full-page). Format is auto-detected by extension and content: .mmd/.md/.mermaid → Mermaid, .drawio → draw.io. Mermaid supports only flowchart/graph, sequenceDiagram, stateDiagram-v2 and erDiagram; other kinds (gantt/journey/mindmap/timeline) raise unsupported diagram kind — use draw.io instead. Existing output is not overwritten by default; overwrite=true allows replacement.

- **Parameters**: `source: str`, `out: str`, `direction: str`, `page: int | str`, `overwrite: bool`
- **Returns**: `dict`
- **Flags**: normal operation

---

### `install_skill`

Install the diagram-design and offipy-diagram skills into the host agent's skill directory (default ~/.claude/skills/, --target_dir to override). Idempotent: skips if the target already exists (never overwrites user edits); --force deletes and rebuilds the target directory.

- **Parameters**: `target_dir: str`, `force: bool`
- **Returns**: `dict`
- **Flags**: normal operation
