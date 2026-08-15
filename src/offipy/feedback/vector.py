"""扁平特征快照 → 定长 numpy 向量（FEATURES schema 对齐）。"""

from __future__ import annotations

import numpy as np

from offipy.art.features_registry import feature_keys


def encode_vector(features: dict[str, float]) -> np.ndarray[tuple[int, ...], np.dtype[np.float64]]:
    """按 feature_keys() 顺序把扁平 dict 编码为 float64 向量；缺失 key 补 0.0。"""
    keys = feature_keys()
    vec = np.zeros(len(keys), dtype=np.float64)
    for i, key in enumerate(keys):
        vec[i] = features.get(key, 0.0)
    return vec
