"""推理：从有效 model.json 推导三个 head。

两个消费点（互不重叠）：
1. learned_adjustments(profile) —— 历史驱动（读 JSONL 快照 infer worth → 按
   rule_id 聚合 → 量化 → feedback_severity_adjustments，契约同 v2 的
   recommend_adjustments；无模型/过期 → None → 调用方回退 v2）
2. analyze 时 per-finding —— severity_shift / quality.score（当前 deck features
   现场算）。analyze.py 在函数内惰性 import 本模块（不拖 numpy）。
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import numpy as np

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.art.feedback import load_records

from .heads import quality_score_from_worth, quantize_delta
from .model import load_model, model_file, model_valid, weights_from_dict
from .pairs import valid_records
from .vector import encode_vector

if TYPE_CHECKING:
    from pathlib import Path

    from .mlp import MLP


def _load_valid_model(feedback_dir: str | Path | None) -> tuple[MLP, dict[str, Any]] | None:
    data = load_model(model_file(feedback_dir))
    if data is None or not model_valid(data, feature_schema_version()):
        return None
    try:
        mlp = weights_from_dict(
            data, input_dim=len(feature_keys()), hidden_dims=tuple(data["hidden_dims"])
        )
    except (ValueError, KeyError, TypeError):
        return None
    return mlp, data


def model_worth(features: dict[str, float], mlp: MLP | None = None) -> float:
    """单个扁平特征快照 → worth 标量。调用方须传有效 mlp。"""
    assert mlp is not None, "model_worth 需要有效 MLP"
    return mlp.predict(encode_vector(features))


def learned_adjustments(
    profile: str, *, feedback_dir: str | Path | None = None
) -> dict[str, int] | None:
    """rule.delta head：历史记录 → {rule_id: ±1}。无有效模型 → None（冷启动 v2）。"""
    loaded = _load_valid_model(feedback_dir)
    if loaded is None:
        return None
    mlp, _ = loaded
    records = valid_records(load_records(feedback_dir), profile=profile)
    agg: dict[str, list[float]] = {}
    for rec in records:
        # rec.features 在 valid_records 过滤后仍是 Optional——`or {}` 收窄且语义无害
        agg.setdefault(rec.rule_id, []).append(model_worth(rec.features or {}, mlp))
    result: dict[str, int] = {}
    for rule_id, worths in agg.items():
        delta = quantize_delta(float(np.mean(worths)))
        if delta != 0:
            result[rule_id] = delta
    return result


def quality_score_for_report(mean_worth: float) -> float:
    """deck 全部 finding 的 worth 均值 → quality.score。"""
    return quality_score_from_worth(mean_worth)
