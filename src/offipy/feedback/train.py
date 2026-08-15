"""离线训练编排：读记录 → 过滤 → 配对 → MLP 训练 → 原子写 model.json。

F2-E：只有成功训练才原子写；样本不足 / 无有效样本时返回状态 JSON，旧模型
保留不删不覆盖。纯 CPU numpy，不依赖 Office。
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import numpy as np

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.art.feedback import load_records

from .mlp import EPOCHS, HIDDEN_DIMS, LR, MARGIN, MLP, REG_WEIGHT, SEED
from .model import model_file, save_model
from .pairs import build_pairs, valid_records
from .registry import OUTPUT_SCHEMA_VERSION
from .vector import encode_vector

_MIN_PAIRS = 50
# #112：判别力门禁阈值——训练后模型对全量样本输出的 worth 标准差低于该值视为
# 坍缩成常数（ReLU 全死 / 退化），是坏模型，拒绝写盘。
MIN_OUTPUT_STD = 1e-6


def run_training(
    feedback_dir: str | Path | None = None,
    *,
    seed: int = SEED,
    min_pairs: int = _MIN_PAIRS,
) -> dict[str, Any]:
    """训练并写 model.json；失败返回状态 JSON（不抛、不覆盖旧模型）。"""
    dir_path = Path(feedback_dir) if feedback_dir else None
    records = load_records(dir_path)
    valid = valid_records(records)
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
    diverged = False
    for _ in range(EPOCHS):
        loss = mlp.train_step(X_fixed, X_accepted, lr=LR, margin=MARGIN, reg_weight=REG_WEIGHT)
        if not math.isfinite(loss):
            diverged = True
            break
    if diverged:
        # #112：数值爆炸（loss=inf/nan）→ 不写模型，返回状态
        return {
            "trained": False,
            "reason": "training_diverged",
            "pairs": len(pairs),
            "samples": len(valid),
            "loss": str(loss),
        }
    # #112：判别力门禁——常数输出（ReLU 全死/退化）是坏模型，拒绝写
    outputs = np.concatenate([mlp.predict_batch(X_fixed), mlp.predict_batch(X_accepted)])
    output_std = float(outputs.std())
    if output_std < MIN_OUTPUT_STD:
        return {
            "trained": False,
            "reason": "model_collapsed",
            "pairs": len(pairs),
            "samples": len(valid),
            "loss": round(loss, 4),
            "output_std": round(output_std, 6),
        }
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
