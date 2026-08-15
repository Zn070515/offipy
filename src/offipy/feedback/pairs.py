"""配对构建器：同 (rule_id, profile) 组内 fixed > accepted 有序对。

ignored 排除；配对数量 = Σ fixed×accepted（同一 fixed 可与多个 accepted 配，
数据利用充分）。纯 python，无 numpy——status.py 顶层可安全 import。
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Iterable

    from offipy.art.feedback import ArtFeedbackRecord


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
