"""扁平特征 dict → 定长 numpy 向量（按 feature_keys 对齐，缺失补 0）。"""

import numpy as np

from offipy.art.features_registry import feature_keys
from offipy.feedback.vector import encode_vector


def test_encode_vector_length_and_alignment():
    keys = feature_keys()
    feats = {keys[0]: 1.5, keys[3]: -2.0}
    vec = encode_vector(feats)
    assert vec.shape == (len(keys),)
    assert vec.dtype == np.float64
    assert vec[0] == 1.5
    assert vec[3] == -2.0
    assert vec[1] == 0.0


def test_encode_vector_empty_dict_zero():
    vec = encode_vector({})
    assert vec.shape == (len(feature_keys()),)
    assert not vec.any()
