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

from .mlp import MLP, SEED, TrainingDiverged, adaptive_hidden_dims, capacity_report, train_mlp
from .model import model_file, save_model
from .pairs import build_pairs, per_rule_diagnosis, record_filter_breakdown, valid_records
from .preprocess import fit_preprocessing, transform_features
from .registry import OUTPUT_SCHEMA_VERSION
from .validation import repeated_stratified_cv
from .vector import encode_vector

_MIN_PAIRS = 50
# #122 A6：ensemble member 数——多 seed 训练取平均降方差，member 之间只 seed 不同。
K = 5
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
        # #131：无有效样本时给出被过滤原因分类，不再静默返回。
        breakdown = record_filter_breakdown(records)
        return {
            "trained": False,
            "reason": "no_valid_samples",
            "samples": len(records),
            "excluded": {k: v for k, v in breakdown.items() if k != "valid"},
            "hint": (
                "无带特征快照的可训练样本：features 只能在 analyze 现场编码，"
                "请经 deck audit 内联标注或 append 时传 --features"
            ),
        }
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
    # #122 A6：ensemble K 个 member（只 seed 不同）——多起点平均降方差。
    # seed 用 run_training 的 seed 参数（不是硬编码 SEED），保持 seed 参数有意义。
    seeds = [seed + i for i in range(K)]
    members: list[tuple[int, MLP]] = []
    try:
        members = [(s, train_mlp(X_fixed, X_accepted, hidden_dims, seed=s)) for s in seeds]
    except TrainingDiverged:
        # #112：数值爆炸（loss=inf/nan）→ 不写模型，返回状态
        return {
            "trained": False,
            "reason": "training_diverged",
            "pairs": len(pairs),
            "samples": len(valid),
        }
    # #112 判别力门禁 + #122 A6 校准：全 member 对全量样本的 worth。
    # member_means = 每样本的 ensemble 均值（axis=1 跨 K member 平均；不是 axis=0 跨
    # fixed/accepted 条件平均——那样对称训练下固定+可接受相互抵消，std≈0 误报坍缩）。
    all_wf = np.array([m.predict_batch(X_fixed) for _, m in members])  # (K, n_pairs, 1)
    all_wa = np.array([m.predict_batch(X_accepted) for _, m in members])
    member_means = np.concatenate([all_wf, all_wa], axis=1).mean(axis=0)  # (2*n_pairs, 1)
    output_std = float(member_means.std())
    if output_std < MIN_OUTPUT_STD:
        return {
            "trained": False,
            "reason": "model_collapsed",
            "pairs": len(pairs),
            "samples": len(valid),
            "output_std": round(output_std, 6),
        }
    # #122 A6：校准 + abstain（#112 gate 之后、写盘之前）。
    # worth_scale 用训练分布的 worth 幅值（std）——quality_score sigmoid 前的归一。
    worth_scale = float(member_means.std()) or 1.0
    # 跨 member 的 |worth| 离散度：member 间分歧大（std 高）→ 该样本不确定 → abstain。
    abs_w = np.abs(np.concatenate([all_wf, all_wa], axis=1))  # (K, 2*n_pairs, 1)
    stds = abs_w.std(axis=0)
    abstain = {
        "worth_margin_p25": float(np.percentile(np.abs(member_means), 25)),
        "std_p80": float(np.percentile(stds, 80)),
    }
    rules_with_pairs = len({f.rule_id for f, _ in pairs})
    stats = {
        "pairs": len(pairs),
        "samples": len(valid),
        "rules_with_pairs": rules_with_pairs,
        "capacity": capacity,
        "ensemble_size": K,
        "calibration": {"worth_scale": worth_scale},
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
        members=members,
        input_schema_version=feature_schema_version(),
        output_schema_version=OUTPUT_SCHEMA_VERSION,
        seed=seed,
        stats=stats,
        preprocessing=pre,
        calibration={"worth_scale": worth_scale},
        abstain=abstain,
        path=path,
    )
    return {
        "trained": True,
        **stats,
        "model": str(path),
        "per_rule": per_rule_diagnosis(valid, min_pairs),
    }
