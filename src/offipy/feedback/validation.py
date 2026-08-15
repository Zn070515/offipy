"""样本级 repeated stratified CV（Kajimura 2022：绝不 pair-level split）。

按 rule_id 对独立样本分层分 5 折；每折在训练样本上重拟合预处理 + 重建 pairs +
训练单 member（hidden_dims 用最终模型的，val 是 proxy）；在留出样本的 pairs 上
评估 pairwise accuracy（worth(fixed)>worth(accepted)，平局 0.5）与 separation。
无 early stopping（Rajabi & Ribeiro 2024）。输出 95% 下界：下界触 0.5 →
与随机不可区分 → random_indistinguishable=True（调用方 soft flag，不 hard reject）。
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, Any

import numpy as np

from .mlp import MLP, TrainingDiverged, train_mlp
from .pairs import build_pairs
from .preprocess import fit_preprocessing, transform_features
from .vector import encode_vector

if TYPE_CHECKING:
    from offipy.art.feedback import ArtFeedbackRecord

FOLDS = 5
REPEATS = 3
# 95% 置信 z 值（近似正态下界）
_Z_95 = 1.96
# 随机不可区分阈值：95% 下界触 chance accuracy 视为与随机不可区分
_CHANCE_ACC = 0.5


def _fold_splits(
    records: list[ArtFeedbackRecord], *, k: int, rng: np.random.Generator
) -> list[list[ArtFeedbackRecord]]:
    """按 rule_id 分层：每 rule 记录 shuffle 后轮流分配 fold（确定性）。"""
    by_rule: dict[str, list[ArtFeedbackRecord]] = {}
    for r in records:
        by_rule.setdefault(r.rule_id, []).append(r)
    folds: list[list[ArtFeedbackRecord]] = [[] for _ in range(k)]
    for group in by_rule.values():
        idx = list(range(len(group)))
        rng.shuffle(idx)
        for pos, gi in enumerate(idx):
            folds[pos % k].append(group[gi])
    return folds


def _eval_pairs(
    mlp: MLP,
    pairs: list[tuple[ArtFeedbackRecord, ArtFeedbackRecord]],
    pre: dict[str, Any],
) -> tuple[float, float]:
    """留出折 pairs 上的 pairwise accuracy（平局 0.5）与 separation（worth 均值差）。"""
    wf = np.array([mlp.predict(transform_features(f.features or {}, pre)) for f, _ in pairs])
    wa = np.array([mlp.predict(transform_features(a.features or {}, pre)) for _, a in pairs])
    acc = float((wf > wa).mean()) + 0.5 * float((wf == wa).mean())
    return acc, float(wf.mean() - wa.mean())


def repeated_stratified_cv(
    records: list[ArtFeedbackRecord],
    *,
    hidden_dims: tuple[int, ...],
    seed: int,
    k: int = FOLDS,
    repeats: int = REPEATS,
) -> dict[str, Any]:
    """重复分层 CV：返回 {accuracy_pooled, accuracy_mean, accuracy_std, separation,
    folds, random_indistinguishable}。

    全部守卫，绝不抛：空折 / 空 train_pairs / 空 val_pairs / 单折 TrainingDiverged
    均跳过该折（CV 是软代理，单折发散不 hard reject）。
    """
    rng = np.random.default_rng(seed)
    per_fold_acc: list[float] = []
    per_fold_sep: list[float] = []
    pooled_acc_sum: float = 0.0
    pooled_n = 0
    for _ in range(repeats):
        folds = _fold_splits(records, k=k, rng=rng)
        for fi in range(k):
            train_recs = [r for i, fold in enumerate(folds) if i != fi for r in fold]
            val_recs = folds[fi]
            if len(train_recs) < 2:
                continue
            Xtr = np.array([encode_vector(r.features or {}) for r in train_recs])
            pre = fit_preprocessing(Xtr)
            train_pairs = build_pairs(train_recs)
            if not train_pairs:
                continue
            Xf = np.array([transform_features(f.features or {}, pre) for f, _ in train_pairs])
            Xa = np.array([transform_features(a.features or {}, pre) for _, a in train_pairs])
            try:
                mlp = train_mlp(Xf, Xa, hidden_dims, seed=seed + fi)
            except TrainingDiverged:
                continue
            val_pairs = build_pairs(val_recs)
            if not val_pairs:
                continue
            acc, sep = _eval_pairs(mlp, val_pairs, pre)
            per_fold_acc.append(acc)
            per_fold_sep.append(sep)
            pooled_acc_sum += acc * len(val_pairs)
            pooled_n += len(val_pairs)
    n_folds = len(per_fold_acc)
    if n_folds == 0:
        return {
            "accuracy_pooled": 0.0,
            "accuracy_mean": 0.0,
            "accuracy_std": 0.0,
            "separation": 0.0,
            "folds": 0,
            "random_indistinguishable": True,
        }
    mean_acc = float(np.mean(per_fold_acc))
    std_acc = float(np.std(per_fold_acc, ddof=1)) if n_folds > 1 else 0.0
    lower95 = mean_acc - _Z_95 * std_acc / math.sqrt(n_folds)
    return {
        "accuracy_pooled": round(pooled_acc_sum / pooled_n, 4) if pooled_n else 0.0,
        "accuracy_mean": round(mean_acc, 4),
        "accuracy_std": round(std_acc, 4),
        "separation": round(float(np.mean(per_fold_sep)), 4),
        "folds": n_folds,
        "random_indistinguishable": lower95 <= _CHANCE_ACC,
    }
