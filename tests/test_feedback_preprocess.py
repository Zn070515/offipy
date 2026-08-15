"""preprocess：零方差 drop + 高相关去重 + z-score 标准化（#118/#120）。"""

import numpy as np

from offipy.art.features_registry import feature_keys
from offipy.feedback.preprocess import fit_preprocessing, transform_features, transform_vector


def test_zero_variance_column_dropped():
    # 第 2 列零方差（全 0）→ 不进 kept；其余两列保留
    X = np.array(
        [
            [1.0, 0.0, 5.0],
            [2.0, 0.0, 1.0],
            [3.0, 0.0, 5.0],
            [4.0, 0.0, 1.0],
        ],
        dtype=np.float64,
    )
    pre = fit_preprocessing(X)
    assert 1 not in pre["kept"]
    assert pre["kept"] == [0, 2]
    assert len(pre["mean"]) == 2
    assert len(pre["scale"]) == 2


def test_perfectly_correlated_columns_dedup_keep_earlier():
    # col1 = 2 * col0（完美相关）；col2 零方差先被 drop → 只留 registry 靠前的 col0
    X = np.array(
        [
            [1.0, 2.0, 3.0],
            [2.0, 4.0, 3.0],
            [3.0, 6.0, 3.0],
        ],
        dtype=np.float64,
    )
    pre = fit_preprocessing(X)
    assert pre["kept"] == [0]


def test_negative_correlation_dedup_keep_earlier():
    # col1 = -2 * col0（完美负相关）→ 同样去重，只留 registry 靠前的 col0
    X = np.array(
        [
            [1.0, -2.0, 3.0],
            [2.0, -4.0, 3.0],
            [3.0, -6.0, 3.0],
        ],
        dtype=np.float64,
    )
    pre = fit_preprocessing(X)
    assert pre["kept"] == [0]


def test_near_but_below_corr_threshold_keeps_columns():
    # |r|=0.98 < 0.99 → 不去重（阈值边界：只 drop |r|>=0.99）
    a_c = np.array([-1.0, 0.0, 1.0])
    t = 0.117233
    b_c = a_c + t * np.array([1.0, -2.0, 1.0])  # w=[1,-2,1] 与 a_c 正交且已中心化
    assert abs(float(np.corrcoef(a_c, b_c)[0, 1])) < 0.99
    X = np.column_stack([a_c + 5.0, b_c, np.zeros(3)]).astype(np.float64)  # 第三列零方差 → drop
    pre = fit_preprocessing(X)
    assert pre["kept"] == [0, 1]


def test_all_constant_columns_fallback_transform_to_zero():
    # 全列零方差 → kept 回退全列、scale=1.0；transform 后常数列 → 0
    X = np.array(
        [
            [5.0, 7.0, 9.0],
            [5.0, 7.0, 9.0],
            [5.0, 7.0, 9.0],
        ],
        dtype=np.float64,
    )
    pre = fit_preprocessing(X)
    assert pre["kept"] == [0, 1, 2]
    assert pre["scale"] == [1.0, 1.0, 1.0]
    out = transform_vector(np.array([5.0, 7.0, 9.0], dtype=np.float64), pre)
    np.testing.assert_array_almost_equal(out, np.zeros(3))


def test_transform_vector_subset_standardizes():
    pre = {"kept": [0, 5], "mean": [1.0, 2.0], "scale": [2.0, 0.5]}
    raw = np.array([3.0, 99.0, 99.0, 99.0, 99.0, 4.0], dtype=np.float64)
    out = transform_vector(raw, pre)
    np.testing.assert_array_almost_equal(out, np.array([1.0, 4.0]))  # (3-1)/2, (4-2)/0.5


def test_transform_features_roundtrip_missing_defaults_zero():
    keys = feature_keys()
    N = len(keys)
    pre = {"kept": [0, N - 1], "mean": [0.0, 0.0], "scale": [1.0, 1.0]}
    out = transform_features({keys[0]: 2.0}, pre)  # keys[N-1] 缺失 → encode_vector 补 0.0
    assert out.shape == (2,)
    assert out[0] == 2.0
    assert out[1] == 0.0


def test_fit_then_transform_standardizes_kept_columns():
    rng = np.random.default_rng(7)
    X = rng.normal(3.0, 2.0, size=(200, 5)).astype(np.float64)
    X[:, 1] = 0.0  # 零方差 → drop
    X[:, 4] = 3.0 * X[:, 0]  # 与 col0 完美相关 → dedup drop
    pre = fit_preprocessing(X)
    Z = np.stack([transform_vector(row, pre) for row in X])
    assert Z.shape[1] == 3  # 5 列 → 去零方差(1) + 去相关(1) → 剩 3
    np.testing.assert_allclose(Z.mean(axis=0), np.zeros(3), atol=1e-10)
    np.testing.assert_allclose(Z.std(axis=0), np.ones(3), atol=1e-10)
