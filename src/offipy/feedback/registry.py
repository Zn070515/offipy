"""OUTPUTS 输出注册表（offipy.feedback）。

每个 head 一条 OutputSpec。v0.17 全 derived（从 worth 推导，不依赖权重形状）——
新增派生 head 不需要重训、不会让旧模型过期。output_schema_version 只记录
（训练时有哪些 head），不门禁推理。若未来有 kind="learned" 的 head（自带权重），
其可用性单独按自身 schema 门禁。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from offipy.art.profiles import ALL_RULES

from .heads import quality_score_from_worth, quantize_delta, severity_shift_from_worth

if TYPE_CHECKING:
    from collections.abc import Callable

OUTPUT_SCHEMA_VERSION = "1"


@dataclass(frozen=True)
class OutputSpec:
    id: str
    kind: str  # "derived" | "learned"
    domain: str
    integration: str
    derivation: Callable[[float], Any] | None = None
    enabled: bool = True


OUTPUTS: dict[str, OutputSpec] = {}
for _rule_id in sorted(ALL_RULES):
    OUTPUTS[f"rule.delta.{_rule_id}"] = OutputSpec(
        id=f"rule.delta.{_rule_id}",
        kind="derived",
        domain="[-2,2]",
        integration="feedback_severity_adjustments → art/rules.py:157（历史驱动）",
        derivation=quantize_delta,
    )
OUTPUTS["finding.severity_shift"] = OutputSpec(
    id="finding.severity_shift",
    kind="derived",
    domain="[-1,1]",
    integration="analyze.py 后处理 pass（severity_override=False 才作用）",
    derivation=severity_shift_from_worth,
)
OUTPUTS["quality.score"] = OutputSpec(
    id="quality.score",
    kind="derived",
    domain="[0,100]",
    integration="替换 analyze.py:53 _experimental_score（opt-in）",
    derivation=quality_score_from_worth,
)
