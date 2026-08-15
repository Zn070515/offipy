"""离线训练编排：读记录 → 过滤 → 配对 → MLP 训练 → 原子写 model.json。

F2-E：只有成功训练才原子写；样本不足 / 无有效样本时返回状态 JSON，旧模型
保留不删不覆盖。纯 CPU numpy，不依赖 Office。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.art.feedback import load_records

from .mlp import EPOCHS, HIDDEN_DIMS, LR, MARGIN, MLP, REG_WEIGHT, SEED
from .model import model_file, save_model
from .pairs import build_pairs
from .registry import OUTPUT_SCHEMA_VERSION
from .vector import encode_vector

_MIN_PAIRS = 50


def run_training(
    feedback_dir: str | Path | None = None,
    *,
    seed: int = SEED,
    min_pairs: int = _MIN_PAIRS,
) -> dict[str, Any]:
    """训练并写 model.json；失败返回状态 JSON（不抛、不覆盖旧模型）。"""
    dir_path = Path(feedback_dir) if feedback_dir else None
    records = load_records(dir_path)
    valid = [
        r
        for r in records
        if r.action in ("fixed", "accepted")
        and r.features is not None
        and r.feature_schema_version == feature_schema_version()
    ]
    if not valid:
        return {"trained": False, "reason": "no_valid_samples", "samples": len(records)}
    pairs = build_pairs(valid)
    if len(pairs) < min_pairs:
        return {
            "trained": False,
            "reason": "insufficient_pairs",
            "pairs": len(pairs),
            "samples": len(valid),
        }
    X_fixed = np.array([encode_vector(f.features or {}) for f, _ in pairs])
    X_accepted = np.array([encode_vector(a.features or {}) for _, a in pairs])
    mlp = MLP(input_dim=len(feature_keys()), hidden_dims=HIDDEN_DIMS, seed=seed)
    loss = 0.0
    for _ in range(EPOCHS):
        loss = mlp.train_step(X_fixed, X_accepted, lr=LR, margin=MARGIN, reg_weight=REG_WEIGHT)
    rules_with_pairs = len({f.rule_id for f, _ in pairs})
    stats = {
        "pairs": len(pairs),
        "samples": len(valid),
        "loss": round(loss, 4),
        "rules_with_pairs": rules_with_pairs,
    }
    path = model_file(dir_path)
    save_model(
        mlp,
        input_schema_version=feature_schema_version(),
        output_schema_version=OUTPUT_SCHEMA_VERSION,
        seed=seed,
        hidden_dims=HIDDEN_DIMS,
        stats=stats,
        path=path,
    )
    return {"trained": True, **stats, "model": str(path)}
