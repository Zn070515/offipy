> [English](index.en.md)

# API 参考

本参考由 `scripts/gen_api_ref.py` 从 `schema.py` 单一来源生成，覆盖 server / CLI / MCP 三入口的同一批操作。

| 应用 | 操作数 | 只读 | 改动状态 |
| --- | --- | --- | --- |
| [Excel](excel.md) | 25 | 4 | 17 |
| [Word](word.md) | 32 | 3 | 25 |
| [PowerPoint](ppt.md) | 17 | 4 | 8 |

每个操作：`doc_id` 缺省走当前活动文档（Excel `bookN` / Word `docN` / PPT `presN`）；`expected_target` 用于破坏性操作的绑定校验。
