"""配对构建器 + 共享 valid_records 过滤：同 (rule_id, profile) 组内 fixed > accepted 有序对。

ignored 排除；配对数量 = Σ fixed×accepted（同一 fixed 可与多个 accepted 配，
数据利用充分）。valid_records 是可训练样本的单一事实来源（feedback.status /
feedback.infer 共用，避免消费方各自手抄过滤逻辑而漂移），可选按 profile 过滤。
纯 python，无 numpy——status.py 顶层可安全 import。
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from offipy.art.features_registry import feature_schema_version

if TYPE_CHECKING:
    from collections.abc import Iterable

    from offipy.art.feedback import ArtFeedbackRecord


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
