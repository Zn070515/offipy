"""配对构建器 + 共享 valid_records 过滤：同 (rule_id, profile) 组内 fixed > accepted 有序对。

ignored 排除；配对数量 = Σ fixed×accepted（同一 fixed 可与多个 accepted 配，
数据利用充分）。valid_records 是可训练样本的单一事实来源（feedback.status /
feedback.infer 共用，避免消费方各自手抄过滤逻辑而漂移），可选按 profile 过滤。
纯 python，无 numpy——status.py 顶层可安全 import。
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING

from offipy.art.features_registry import feature_schema_version

if TYPE_CHECKING:
    from collections.abc import Iterable

    from offipy.art.feedback import ArtFeedbackRecord


# 默认最小配对门限：train 的 min_pairs 默认值与 status 的逐规则诊断视野共用，
# 放此处避免跨模块重复字面量漂移（pairs 纯 python 无 numpy，status 顶层可安全 import）。
MIN_PAIRS = 50


def valid_records(
    records: Iterable[ArtFeedbackRecord],
    profile: str | None = None,
) -> list[ArtFeedbackRecord]:
    """可训练样本：fixed/accepted、带 features、schema 与当前一致；可选按 profile 过滤。"""
    return [
        r
        for r in records
        if r.action in ("fixed", "accepted")
        and r.features is not None
        and r.feature_schema_version == feature_schema_version()
        and (profile is None or r.profile == profile)
    ]


def build_pairs(
    records: Iterable[ArtFeedbackRecord],
) -> list[tuple[ArtFeedbackRecord, ArtFeedbackRecord]]:
    groups: dict[tuple[str, str], list[ArtFeedbackRecord]] = {}
    for rec in records:
        if rec.action == "ignored":
            continue
        groups.setdefault((rec.rule_id, rec.profile), []).append(rec)
    pairs: list[tuple[ArtFeedbackRecord, ArtFeedbackRecord]] = []
    for group in groups.values():
        fixed = [r for r in group if r.action == "fixed"]
        accepted = [r for r in group if r.action == "accepted"]
        pairs.extend((f, a) for f in fixed for a in accepted)
    return pairs


PER_RULE_MIN_BASE = 10


def per_rule_diagnosis(
    records: Iterable[ArtFeedbackRecord],
    min_pairs: int,
) -> dict[str, dict[str, int | bool]]:
    """逐规则样本诊断：{fixed, accepted, pairs, single_direction, suggest}。

    供 insufficient_pairs 分支给用户可行动建议。内部按 (rule_id, profile)
    分组（与 build_pairs 一致），配对语义不虚报：跨 profile 的 fixed/accepted
    不配对。每规则输出规则级总量 fixed/accepted、实际配对 pairs=Σ fixed_p×accepted_p
    （== build_pairs 计数）、single_direction=(pairs==0)。per_rule_min =
    max(ceil(min_pairs/n_rules), 10)；deficit 规则缺的样本数按
    ceil(deficit/max(fixed,accepted)) 给 suggest（单方向时补另一侧）。"""
    # (rule_id, profile) -> {fixed, accepted}；与 build_pairs 同键，配对语义一致
    groups: dict[str, dict[str, dict[str, int]]] = {}
    for rec in records:
        if rec.action == "ignored":
            continue
        g = groups.setdefault(rec.rule_id, {}).setdefault(rec.profile, {"fixed": 0, "accepted": 0})
        g[rec.action] += 1
    if not groups:
        return {}
    per_rule_min = max(math.ceil(min_pairs / len(groups)), PER_RULE_MIN_BASE)
    out: dict[str, dict[str, int | bool]] = {}
    for rule_id, profiles in groups.items():
        fixed = sum(g["fixed"] for g in profiles.values())
        accepted = sum(g["accepted"] for g in profiles.values())
        pairs = sum(g["fixed"] * g["accepted"] for g in profiles.values())
        single_direction = pairs == 0
        suggest = 0
        if pairs < per_rule_min:
            denom = max(fixed, accepted)
            if denom:
                suggest = math.ceil((per_rule_min - pairs) / denom)
        out[rule_id] = {
            "fixed": fixed,
            "accepted": accepted,
            "pairs": pairs,
            "single_direction": single_direction,
            "suggest": suggest,
        }
    return out


def record_filter_breakdown(
    records: Iterable[ArtFeedbackRecord],
) -> dict[str, int]:
    """记录过滤分类：valid / no_features / schema_mismatch / ignored / other。

    供 status（#144 excluded 明细）与 train（#131 未采样提示）共用，让「被静默
    过滤」变成「显式分类可见」。valid 判定与 valid_records 同判据（fixed/accepted
    + features 非 None + schema 匹配）——若日后修改 valid_records 判据，须同步此处。
    """
    current_schema = feature_schema_version()
    out = {"valid": 0, "no_features": 0, "schema_mismatch": 0, "ignored": 0, "other": 0}
    for r in records:
        if r.action == "ignored":
            out["ignored"] += 1
        elif r.action not in ("fixed", "accepted"):
            out["other"] += 1
        elif r.features is None:
            out["no_features"] += 1
        elif r.feature_schema_version != current_schema:
            out["schema_mismatch"] += 1
        else:
            out["valid"] += 1
    return out
