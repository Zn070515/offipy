"""离线训练编排：读记录 → 过滤 → 配对 → MLP 训练 → 原子写 model.json。

F2-E：只有成功训练才原子写；样本不足 / 无有效样本时返回状态 JSON，旧模型
保留不删不覆盖。纯 CPU numpy，不依赖 Office。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np

from offipy.art.features_registry import feature_schema_version
from offipy.art.feedback import load_records

from .mlp import SEED, TrainingDiverged, adaptive_hidden_dims, capacity_report, train_mlp
from .model import model_file, save_model
from .pairs import build_pairs, per_rule_diagnosis, valid_records
from .preprocess import fit_preprocessing, transform_features
from .registry import OUTPUT_SCHEMA_VERSION
from .validation import repeated_stratified_cv
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
    # #112：NaN 特征 → 不静默 drop 也不训练，保持契约（training_diverged 不写模型）
    X_raw = np.array([encode_vector(r.features or {}) for r in valid])
    if not np.isfinite(X_raw).all():
        return {"trained": False, "reason": "training_diverged", "samples": len(valid)}
    pairs = build_pairs(valid)
    # 空 pairs（全 fixed 无 accepted / 反之）即使 min_pairs=0 也属样本不足：
    # 后续 train_mlp 依赖 X_fixed.shape[1]，空数组会 IndexError，违反不抛契约。
    if not pairs or len(pairs) < min_pairs:
        return {
            "trained": False,
            "reason": "insufficient_pairs",
            "pairs": len(pairs),
            "samples": len(valid),
            "per_rule": per_rule_diagnosis(valid, min_pairs),
        }
    # 预处理在独立样本上拟合（零方差 drop + 高相关去重 + z-score），配对与推理共享
    pre = fit_preprocessing(X_raw)
    X_fixed = np.array([transform_features(f.features or {}, pre) for f, _ in pairs])
    X_accepted = np.array([transform_features(a.features or {}, pre) for _, a in pairs])
    input_dim = len(pre["kept"])
    # #121 A3：容量按样本数自适应（Lunt & Xu H≈√n，[4,32]，n≥120 双层）。
    # 独立样本数 n（不是 pairs）驱动容量；capacity 是软告警，只记录不拒绝写盘。
    n = len(valid)
    hidden_dims = adaptive_hidden_dims(n)
    capacity = capacity_report(n, input_dim, hidden_dims)
    try:
        mlp = train_mlp(X_fixed, X_accepted, hidden_dims, seed=seed)
    except TrainingDiverged:
        # #112：数值爆炸（loss=inf/nan）→ 不写模型，返回状态
        return {
            "trained": False,
            "reason": "training_diverged",
            "pairs": len(pairs),
            "samples": len(valid),
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
            "output_std": round(output_std, 6),
        }
    rules_with_pairs = len({f.rule_id for f, _ in pairs})
    stats = {
        "pairs": len(pairs),
        "samples": len(valid),
        "rules_with_pairs": rules_with_pairs,
        "capacity": capacity,
    }
    if capacity["level"] != "ok":
        stats["capacity_warning"] = True  # soft：只记录，不拒绝写盘
    # #119 A4：样本级 repeated stratified CV（#112 gate 之后、写盘之前）。
    # 全守卫不抛；poor_generalization 是 soft 记录，不拒绝写盘。
    val = repeated_stratified_cv(valid, hidden_dims=hidden_dims, seed=seed)
    poor_generalization = val["random_indistinguishable"]
    stats["validation"] = val
    stats["poor_generalization"] = poor_generalization
    path = model_file(dir_path)
    save_model(
        members=[(seed, mlp)],
        input_schema_version=feature_schema_version(),
        output_schema_version=OUTPUT_SCHEMA_VERSION,
        seed=seed,
        stats=stats,
        preprocessing=pre,
        calibration={"worth_scale": 1.0},  # A6: 填真值
        abstain={"worth_margin_p25": 0.0, "std_p80": 0.0},  # A6: 填真值
        path=path,
    )
    return {"trained": True, **stats, "model": str(path)}
