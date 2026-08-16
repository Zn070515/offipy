> [English](index.en.md)

# API 参考

本参考由 `scripts/gen_api_ref.py` 从 `schema.py` 单一来源生成，覆盖 server / CLI / MCP 三入口的同一批操作。

| 应用 | 操作数 | 只读 | 改动状态 |
| --- | --- | --- | --- |
| [Excel](excel.md) | 25 | 4 | 17 |
| [Word](word.md) | 32 | 3 | 25 |
| [PowerPoint](ppt.md) | 27 | 5 | 17 |
| [图表](diagram.md) | 2 | 0 | 0 |
| [反馈学习](feedback.md) | 5 | 2 | 0 |

每个操作：`doc_id` 缺省走当前活动文档（Excel `book<hex>` / Word `doc<hex>` / PPT `pres<hex>`，高熵不透明，不可枚举）；`expected_target` 用于破坏性操作的绑定校验。`diagram`/`feedback` 不经过 COM、无 Office 文档目标，不适用 `doc_id`/`expected_target`。

> 静态几何质量门禁与基线回归不经过 `schema.py`（纯解析、无 Office/无 COM），另见 [PPTX 质量审计](../audit.md) 与 [基线回归](../audit-baseline.md)。
