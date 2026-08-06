> [中文](index.md)

# API Reference

This reference is generated from the single source of truth `schema.py` by `scripts/gen_api_ref.py` and covers the same set of operations across the three entry points (server / CLI / MCP).

| App | Operations | Read-only | Mutating |
| --- | --- | --- | --- |
| [Excel](excel.en.md) | 25 | 4 | 17 |
| [Word](word.en.md) | 32 | 3 | 25 |
| [PowerPoint](ppt.en.md) | 18 | 4 | 9 |

Every operation: `doc_id` defaults to the current active document (Excel `bookN` / Word `docN` / PPT `presN`); `expected_target` provides target binding for destructive operations.

> Static geometry quality gates and baseline regression do not go through `schema.py` (pure parsing, no Office/COM); see [PPTX Quality Audit](audit.en.md) and [Baseline Regression](audit-baseline.en.md).
