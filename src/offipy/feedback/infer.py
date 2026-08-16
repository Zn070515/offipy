"""推理：从有效 model.json 推导三个 head（v2 ModelBundle 统一封装）。

两个消费点（互不重叠）：
1. learned_adjustments(profile) —— 历史驱动（读 JSONL 快照 infer worth → 按
   rule_id 聚合 → 量化 → feedback_severity_adjustments，契约同 v2 的
   recommend_adjustments；无模型/过期 → None → 调用方回退 v2）
2. analyze 时 per-finding —— severity_shift / quality.score（当前 deck features
   现场算）。analyze.py 在函数内惰性 import 本模块（不拖 numpy）。

ModelBundle 是 A6 推理抽象：ensemble members + preprocessing + calibration +
abstain 一次 load 封装，消费方只拿 bundle 走 worth_mean / should_abstain /
ood_flagged / quality_score。monkeypatch seam：module 级 model_worth 保持
(features, bundle) 两位置参签名，测试直接 patch 它。
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import numpy as np

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.art.feedback import load_records

from .heads import quality_score_from_worth, quantize_delta
from .model import load_model, model_file, model_valid, weights_from_dict
from .pairs import valid_records
from .preprocess import transform_features

if TYPE_CHECKING:
    from pathlib import Path

    from .mlp import MLP


class ModelBundle:
    """v2 model.json 推理封装：ensemble members + preprocessing + calibration + abstain。"""

    def __init__(
        self,
        members: list[MLP],
        pre: dict[str, Any],
        calibration: dict[str, Any],
        abstain: dict[str, Any],
    ) -> None:
        self._members = members
        self._pre = pre
        self._cal = calibration
        self._abs = abstain

    def worth_mean(self, features: dict[str, float]) -> float:
        x = transform_features(features, self._pre)
        vals = [m.predict(x) for m in self._members]
        return float(np.mean(vals))

    def worth_stats(self, features: dict[str, float]) -> tuple[float, float]:
        """(ensemble 均值, member 分歧 std)——一次前向同时给 mean 与 std。

        should_abstain 的 std_p80 分歧分支消费本方法；同时保留为 bundle 公共
        API（test 直测均值/方差与 member 分歧）。
        """
        x = transform_features(features, self._pre)
        vals = np.array([m.predict(x) for m in self._members])
        return float(vals.mean()), float(vals.std())

    def quality_score(self, mean_worth: float) -> float:
        return quality_score_from_worth(mean_worth, self._cal.get("worth_scale", 1.0))

    def should_abstain(self, features: dict[str, float]) -> bool:
        """|worth| < margin_p25（近零不确定）或 member 分歧 std > std_p80 → 不 shift。

        std_p80 缺失时用 inf → 恒不触发，兼容无该键的旧模型。
        """
        mean_w, std_w = self.worth_stats(features)
        margin = float(self._abs.get("worth_margin_p25", 0.0))
        std_p80 = float(self._abs.get("std_p80", float("inf")))
        return abs(mean_w) < margin or std_w > std_p80

    def ood_flagged(self, features: dict[str, float]) -> bool:
        """per-feature z 越界：>30% 特征 |z|>3 或任一 |z|>5 → OOD。"""
        x = transform_features(features, self._pre)
        frac = float((np.abs(x) > 3.0).mean())
        any5 = bool(np.any(np.abs(x) > 5.0))
        return frac > 0.3 or any5

    @classmethod
    def load(cls, feedback_dir: str | Path | None) -> ModelBundle | None:
        data = load_model(model_file(feedback_dir))
        if data is None or not model_valid(data, feature_schema_version()):
            return None
        # #150：schema bump 后 persisted kept 下标可能越过当前 feature_keys——load 阶段
        # 校验，越界/缺失视为无模型回退 v2，杜绝 analyze_scene 抛 IndexError。
        kept = data.get("preprocessing", {}).get("kept")
        if not isinstance(kept, list) or not kept or max(kept) >= len(feature_keys()):
            return None
        if not data.get("members"):
            return None  # 防御：空 ensemble（np.mean([]) → nan）
        try:
            members = [
                weights_from_dict(
                    m,
                    input_dim=len(data["preprocessing"]["kept"]),
                    hidden_dims=tuple(m["hidden_dims"]),
                )
                for m in data["members"]
            ]
        except (ValueError, KeyError, TypeError):
            return None  # 损坏模型（schema 匹配但权重形状错）→ 视为无模型，回退 v2
        return cls(
            members,
            data["preprocessing"],
            data.get("calibration", {}),
            data.get("abstain", {}),
        )


def model_worth(features: dict[str, float], bundle: ModelBundle | None = None) -> float:
    """单个扁平特征快照 → ensemble mean worth。调用方须传有效 bundle。"""
    assert bundle is not None, "model_worth 需要有效 ModelBundle"
    return bundle.worth_mean(features)


def learned_adjustments(
    profile: str, *, feedback_dir: str | Path | None = None
) -> dict[str, int] | None:
    """rule.delta head：历史记录 → {rule_id: ±1}。

    返回值契约（analyze.py 消费方必须区分）：
    - None：无有效模型（缺失/过期/损坏）→ 冷启动回退 v2 的 recommend_adjustments。
    - {}：有有效模型但该 profile 无量化 delta（无记录，或均值量化后全为 0）→
      模型判断「无调整」，应视为权威空结果，不回退 v2。
    """
    if not feedback_dir:
        return None  # #113：无 feedback_dir 不碰全局模型，回退 v2
    bundle = ModelBundle.load(feedback_dir)
    if bundle is None:
        return None
    records = valid_records(load_records(feedback_dir), profile=profile)
    agg: dict[str, list[float]] = {}
    for rec in records:
        # rec.features 在 valid_records 过滤后仍是 Optional——`or {}` 收窄且语义无害
        agg.setdefault(rec.rule_id, []).append(model_worth(rec.features or {}, bundle))
    result: dict[str, int] = {}
    for rule_id, worths in agg.items():
        delta = quantize_delta(float(np.mean(worths)))
        if delta != 0:
            result[rule_id] = delta
    return result


def quality_score_for_report(mean_worth: float, worth_scale: float = 1.0) -> float:
    """通过证据门禁（#111）的 finding 的 worth 均值 → quality.score（薄转发）。"""
    return quality_score_from_worth(mean_worth, worth_scale)
