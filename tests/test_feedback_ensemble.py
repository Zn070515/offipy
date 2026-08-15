"""A6 ModelBundle 推理封装：ensemble mean/stats + calibration + abstain + OOD（#122）。"""

import numpy as np

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.feedback.infer import ModelBundle
from offipy.feedback.mlp import MLP, train_mlp
from offipy.feedback.model import model_file, save_model


def _identity_pre():
    n = len(feature_keys())
    return {"kept": list(range(n)), "mean": [0.0] * n, "scale": [1.0] * n}


def _encode(feats):
    """identity pre 下与 bundle.worth_mean 的 transform 等价（encode_vector 全键 0 缺省）。"""
    return np.array([feats.get(k, 0.0) for k in feature_keys()])


def _write_bundle(tmp_path, *, members=None, calibration=None, abstain=None):
    n = len(feature_keys())
    return save_model(
        members=members or [(0, MLP(input_dim=n, hidden_dims=(4,), seed=0))],
        input_schema_version=feature_schema_version(),
        output_schema_version="1",
        seed=0,
        stats={},
        preprocessing=_identity_pre(),
        calibration=calibration or {"worth_scale": 1.0},
        # 缺省 std_p80=1e9：member 分歧 abstain 恒不触发（真实 std 远小于 1e9），
        # 让 should_abstain 测试专注 margin 路径；std-abstain 测试显式覆盖 std 路径。
        abstain=abstain or {"worth_margin_p25": 0.0, "std_p80": 1e9},
        path=model_file(tmp_path),
    )


def test_load_rebuilds_bundle_from_v2_model(tmp_path):
    n = len(feature_keys())
    members = [
        (0, MLP(input_dim=n, hidden_dims=(4,), seed=0)),
        (1, MLP(input_dim=n, hidden_dims=(4,), seed=1)),
        (2, MLP(input_dim=n, hidden_dims=(4,), seed=2)),
    ]
    _write_bundle(tmp_path, members=members)
    bundle = ModelBundle.load(tmp_path)
    assert bundle is not None
    feats = {"finding.confidence": 0.5}
    expected = np.mean([m.predict(_encode(feats)) for _, m in members])
    assert np.isclose(bundle.worth_mean(feats), float(expected))


def test_worth_stats_matches_member_predictions(tmp_path):
    n = len(feature_keys())
    members = [
        (0, MLP(input_dim=n, hidden_dims=(4,), seed=0)),
        (1, MLP(input_dim=n, hidden_dims=(4,), seed=1)),
    ]
    _write_bundle(tmp_path, members=members)
    bundle = ModelBundle.load(tmp_path)
    assert bundle is not None
    feats = {"finding.confidence": 0.5}
    vals = np.array([m.predict(_encode(feats)) for _, m in members])
    mean, std = bundle.worth_stats(feats)
    assert np.isclose(mean, float(vals.mean()))
    assert np.isclose(std, float(vals.std()))


def test_should_abstain_false_with_zero_margin(tmp_path):
    # 缺省 abstain（margin_p25=0.0、std_p80=1e9）→ margin 与 std 分支都不触发 → 不 abstain
    _write_bundle(tmp_path)
    bundle = ModelBundle.load(tmp_path)
    assert bundle is not None
    assert bundle.should_abstain({"finding.confidence": 0.5}) is False


def test_should_abstain_true_when_members_diverge(tmp_path):
    """std_p80=0 → member 分歧 std > 0 → abstain（分歧路径；margin 分支隔离）。"""
    n = len(feature_keys())
    # 不同 seed 的随机 MLP：权重差异巨大 → 同一输入下输出分歧明显
    members = [
        (0, MLP(input_dim=n, hidden_dims=(4,), seed=0)),
        (1, MLP(input_dim=n, hidden_dims=(4,), seed=1)),
        (2, MLP(input_dim=n, hidden_dims=(4,), seed=2)),
    ]
    _write_bundle(tmp_path, members=members, abstain={"worth_margin_p25": 0.0, "std_p80": 0.0})
    bundle = ModelBundle.load(tmp_path)
    assert bundle is not None
    feats = {"finding.confidence": 0.5}
    _mean, std = bundle.worth_stats(feats)
    assert std > 0.0  # 成员输出确已分歧（前置条件，否则断言无意义）
    assert bundle.should_abstain(feats) is True  # std 分支触发
    # 同 seed 拷贝（无分歧）→ std=0 → 不 abstain（证明是分歧触发，不是 margin）
    dup_members = [(i, MLP(input_dim=n, hidden_dims=(4,), seed=7)) for i in range(3)]
    _write_bundle(tmp_path, members=dup_members, abstain={"worth_margin_p25": 0.0, "std_p80": 0.0})
    dup = ModelBundle.load(tmp_path)
    assert dup is not None
    assert dup.should_abstain(feats) is False


def test_ood_flagged_detects_outlier_feature(tmp_path):
    _write_bundle(tmp_path)
    bundle = ModelBundle.load(tmp_path)
    assert bundle is not None
    # identity pre：x == 原始特征值；confidence=10 → |z|>5 → OOD
    assert bundle.ood_flagged({"finding.confidence": 10.0}) is True
    # 正常特征：全 0 → 无 |z|>3 → 非 OOD
    assert bundle.ood_flagged({}) is False


def test_bundle_is_ensemble_of_distinct_members(tmp_path):
    """真实 train_mlp 训出的 K 个 member：权重/输出确实不同（不是同一函数拷贝）。"""
    rng = np.random.default_rng(7)
    n_fixed, n_acc = 24, 24
    # 可分离数据：fixed x1∈[1,2]、accepted x1∈[-2,-1]，x2 纯噪声
    fixed = np.column_stack([rng.uniform(1.0, 2.0, n_fixed), rng.normal(0, 1, n_fixed)])
    acc = np.column_stack([rng.uniform(-2.0, -1.0, n_acc), rng.normal(0, 1, n_acc)])
    X_fixed = np.array([[0.0] * 4] * n_fixed)  # 4 维占位，训练后只取前 2 维有意义
    X_acc = np.array([[0.0] * 4] * n_acc)
    X_fixed[:, :2] = fixed
    X_acc[:, :2] = acc
    members = [train_mlp(X_fixed, X_acc, (4,), seed=42 + i) for i in range(3)]
    x0 = np.array([1.5, 0.0, 0.0, 0.0])
    preds = [m.predict(x0) for m in members]
    assert len({round(p, 6) for p in preds}) > 1  # member 输出有差异 → ensemble 不是单拷贝
    # 同 seed 则输出一致（确定性校验：差异确实来自 seed 多样性）
    dup = [train_mlp(X_fixed, X_acc, (4,), seed=99)] * 3
    dup_preds = [m.predict(x0) for m in dup]
    assert len({round(p, 6) for p in dup_preds}) == 1
