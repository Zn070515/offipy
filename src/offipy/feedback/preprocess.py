"""预处理：维度筛选（零方差 drop + 高相关去重）+ 全局 z-score 标准化。

在训练集上拟合，mean/scale/kept 持久化到 model.json 的 preprocessing 块；
推理端用同一 transform（transform_features）。missing_default=0.0 语义保留在
encode_vector（缺失补 0），标准化只对 kept 列做。
"""

from __future__ import annotations

from typing import Any

import numpy as np

from .vector import encode_vector

# mypy strict：裸 np.ndarray 触发 type-arg（同 vector.py 的参数化约定）
_Arr = np.ndarray[tuple[int, ...], np.dtype[np.float64]]

_VARIANCE_EPS = 1e-9
_CORR_DEDUP = 0.99


def fit_preprocessing(X: _Arr) -> dict[str, Any]:
    """X: (n, n_features) float64 原始编码向量。返回 {kept, mean, scale}。

    kept 是 feature_keys() 的绝对下标；mean/scale 与 kept 等长。
    scale = std（std>eps）否则 1.0（保 missing/常数列可过）。
    边界：全部列零方差 → kept 回退为全列、scale=1.0（避免 input_dim=0，
    且常数特征标准化为 0，靠 #112 判别力门禁拒绝坍缩模型）。
    """
    f = X.shape[1]
    std = X.std(axis=0)
    keep = np.where(std > _VARIANCE_EPS)[0]
    kept_idx = [int(k) for k in keep]
    if len(kept_idx) > 1:
        # 高相关去重：两两 |r|>=0.99 只留 registry 顺序靠前者
        sub = X[:, keep]
        sub_centered = sub - sub.mean(axis=0)
        norm = np.sqrt((sub_centered**2).sum(axis=0))
        norm[norm == 0] = 1.0
        normed = sub_centered / norm
        drop = set()
        for i in range(normed.shape[1]):
            if i in drop:
                continue
            for j in range(i + 1, normed.shape[1]):
                if j in drop:
                    continue
                if abs(float(normed[:, i] @ normed[:, j])) >= _CORR_DEDUP:
                    drop.add(j)
        kept_idx = [int(keep[i]) for i in range(len(keep)) if i not in drop]
    if not kept_idx:
        kept_idx = list(range(f))
        mean = X.mean(axis=0).tolist()
        scale = [1.0] * f
    else:
        cols = X[:, np.asarray(kept_idx, dtype=np.intp)]
        col_std = cols.std(axis=0)
        mean = cols.mean(axis=0).tolist()
        scale = [float(s) if s > _VARIANCE_EPS else 1.0 for s in col_std]
    return {"kept": kept_idx, "mean": mean, "scale": scale}


def transform_vector(raw: _Arr, pre: dict[str, Any]) -> _Arr:
    """原始全维向量 → 标准化后的 kept 子向量。"""
    kept = np.asarray(pre["kept"], dtype=np.intp)
    x = np.asarray(raw[kept], dtype=np.float64)
    mean = np.asarray(pre["mean"], dtype=np.float64)
    scale = np.asarray(pre["scale"], dtype=np.float64)
    return (x - mean) / scale


def transform_features(features: dict[str, float], pre: dict[str, Any]) -> _Arr:
    """扁平特征 dict → 标准化向量（encode_vector + transform_vector）。"""
    return transform_vector(encode_vector(features), pre)
