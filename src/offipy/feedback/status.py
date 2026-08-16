"""feedback 状态报告：样本 / 配对潜力 / 模型状态。

顶层 numpy-free（只用 stdlib + offipy.art + 本包纯 python 模块），所以 base
install 也能跑 `offipy feedback status`。model 读取走 model.load_model——
它在损坏时返回 None，不抛。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.art.feedback import load_records

from .model import kept_valid, load_model, model_file, model_valid, weights_probe
from .pairs import (
    MIN_PAIRS,
    build_pairs,
    per_rule_diagnosis,
    record_filter_breakdown,
    valid_records,
)


def report_status(feedback_dir: str | Path | None = None) -> dict[str, Any]:
    dir_path = Path(feedback_dir) if feedback_dir else None
    records = load_records(dir_path)
    valid = valid_records(records)
    pairs = build_pairs(valid)
    excluded = {k: v for k, v in record_filter_breakdown(records).items() if k != "valid"}
    # #152：逐规则样本诊断（与 build_pairs 同判据——用 valid 记录，pairs 数与
    # pair_potential 一致）。诊断视野 min_pairs 用共享的 pairs.MIN_PAIRS（与 train
    # 默认同源，避免字面量漂移）；status 顶层 numpy-free，从 pairs import 而非 train。
    per_rule = per_rule_diagnosis(valid, MIN_PAIRS)
    data = load_model(model_file(dir_path))
    if data is not None and model_valid(data, feature_schema_version()):
        # #150：schema 匹配但 kept 越界/缺失/非数值（bump 忘重训）→ stale，不冒充 valid。
        if not kept_valid(data.get("preprocessing", {}), len(feature_keys())):
            return {
                "samples": len(records),
                "valid_samples": len(valid),
                "pair_potential": len(pairs),
                "model": "stale",
                "excluded": excluded,
                "per_rule": per_rule,
            }
        # #147：schema 匹配但权重形状损坏 → corrupt（valid 前最后一道门，不冒充有效模型）。
        if not weights_probe(data):
            return {
                "samples": len(records),
                "valid_samples": len(valid),
                "pair_potential": len(pairs),
                "model": "corrupt",
                "excluded": excluded,
                "per_rule": per_rule,
            }
        pre = data.get("preprocessing", {})
        stats = data.get("stats", {})
        kept = pre.get("kept")
        capacity = stats.get("capacity")
        return {
            "samples": len(records),
            "valid_samples": len(valid),
            "pair_potential": len(pairs),
            "model": "valid",
            "effective_dims": len(kept),
            "samples_per_param": (
                capacity.get("samples_per_param") if isinstance(capacity, dict) else None
            ),
            "poor_generalization": stats.get("poor_generalization"),
            "saturation": stats.get("saturation"),
            "excluded": excluded,
            "per_rule": per_rule,
        }
    return {
        "samples": len(records),
        "valid_samples": len(valid),
        "pair_potential": len(pairs),
        "model": "expired" if data is not None else "none",
        "excluded": excluded,
        "per_rule": per_rule,
    }
