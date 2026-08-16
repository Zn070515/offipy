> [中文](index.md)

# API Reference

This reference is generated from the single source of truth `schema.py` by `scripts/gen_api_ref.py` and covers the same set of operations across the three entry points (server / CLI / MCP).

| App | Operations | Read-only | Mutating |
| --- | --- | --- | --- |
| [Excel](excel.en.md) | 25 | 4 | 17 |
| [Word](word.en.md) | 32 | 3 | 25 |
| [PowerPoint](ppt.en.md) | 27 | 5 | 17 |
| [Diagram](diagram.en.md) | 2 | 0 | 0 |
| [Feedback](feedback.en.md) | 5 | 2 | 0 |

Every operation on the three COM apps (Excel/Word/PowerPoint): `doc_id` defaults to the current active document (Excel `book<hex>` / Word `doc<hex>` / PPT `pres<hex>`, high-entropy and opaque, not enumerable); `expected_target` provides target binding for destructive operations. `diagram`/`feedback` do not go through COM and have no Office document target.

> Static geometry quality gates and baseline regression do not go through `schema.py` (pure parsing, no Office/COM); see [PPTX Quality Audit](../audit.en.md) and [Baseline Regression](../audit-baseline.en.md).
