"""feedback 状态报告：样本 / 配对潜力 / 模型状态。

顶层 numpy-free（只用 stdlib + offipy.art + 本包纯 python 模块），所以 base
install 也能跑 `offipy feedback status`。model 读取走 model.load_model——
它在损坏时返回 None，不抛。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from offipy.art.features_registry import feature_schema_version
from offipy.art.feedback import load_records

from .model import load_model, model_file, model_valid
from .pairs import build_pairs, valid_records


def report_status(feedback_dir: str | Path | None = None) -> dict[str, Any]:
    dir_path = Path(feedback_dir) if feedback_dir else None
    records = load_records(dir_path)
    valid = valid_records(records)
    pairs = build_pairs(valid)
    data = load_model(model_file(dir_path))
    if data is not None and model_valid(data, feature_schema_version()):
        pre = data.get("preprocessing", {})
        stats = data.get("stats", {})
        kept = pre.get("kept")
        return {
            "samples": len(records),
            "valid_samples": len(valid),
            "pair_potential": len(pairs),
            "model": "valid",
            "effective_dims": len(kept) if isinstance(kept, list) else None,
            "samples_per_param": stats.get("capacity", {}).get("samples_per_param"),
            "poor_generalization": stats.get("poor_generalization"),
        }
    return {
        "samples": len(records),
        "valid_samples": len(valid),
        "pair_potential": len(pairs),
        "model": "expired" if data is not None else "none",
    }
