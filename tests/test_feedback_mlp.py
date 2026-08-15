"""MLP：前向形状 / 数值梯度校验 / 配对 margin 单调性 / 确定性。"""

import numpy as np

from offipy.feedback.mlp import MARGIN, MLP, REG_WEIGHT, SEED


def _pairwise_loss(mlp, X_fixed, X_accepted, margin=MARGIN, reg_weight=REG_WEIGHT):
    wf = mlp._predict_batch(X_fixed)
    wa = mlp._predict_batch(X_accepted)
    pair = np.maximum(margin - (wf - wa), 0.0).mean()
    reg = ((wf - 1.0) ** 2).mean() + ((wa + 1.0) ** 2).mean()
    return float(pair + reg_weight * reg)


def test_forward_output_shape():
    mlp = MLP(input_dim=5, hidden_dims=(4, 3))
    X = np.zeros((2, 5))
    out = mlp.forward(X)
    assert out.shape == (2, 1)
    assert mlp.predict(np.zeros(5)) == out[0, 0]


def test_same_seed_same_init():
    a = MLP(input_dim=5, hidden_dims=(4,), seed=SEED)
    b = MLP(input_dim=5, hidden_dims=(4,), seed=SEED)
    for wa, wb in zip(a.W, b.W, strict=True):
        np.testing.assert_array_equal(wa, wb)


def test_different_seed_different_init():
    a = MLP(input_dim=5, hidden_dims=(4,), seed=1)
    b = MLP(input_dim=5, hidden_dims=(4,), seed=2)
    assert not all(np.array_equal(wa, wb) for wa, wb in zip(a.W, b.W, strict=True))


def test_pairwise_loss_decreases_when_fixed_higher():
    mlp = MLP(input_dim=4, hidden_dims=(8,), seed=SEED)
    hi = np.array([[2.0, 0.0, 1.0, 0.5]])
    lo = np.array([[0.5, 0.0, 1.0, 0.5]])
    # 强制第一维权重，让 hi 输出（16）高于 lo（4）：margin 满足 → 配对项 0
    mlp.W[0][0, :] = 1.0
    mlp.W[1][:] = 1.0
    mlp.b[0][:] = 0.0
    mlp.b[1][:] = 0.0
    # 单调性：fixed>accepted 方向的 loss 严格低于反向（reg 项随 worth 幅值变大，
    # 反向配对多出 margin 惩罚，必然更高）
    loss_correct = _pairwise_loss(mlp, hi, lo)
    loss_swapped = _pairwise_loss(mlp, lo, hi)
    assert loss_correct < loss_swapped
    # margin 已满足 → 配对项为 0（reg 项仍 >0，不在此断言）
    pair_term = np.maximum(MARGIN - (mlp._predict_batch(hi) - mlp._predict_batch(lo)), 0.0).mean()
    assert float(pair_term) == 0.0


def test_train_step_reduces_loss():
    rng = np.random.default_rng(0)
    mlp = MLP(input_dim=6, hidden_dims=(12,), seed=SEED)
    X_fixed = rng.normal(size=(4, 6)) + 1.0
    X_accepted = rng.normal(size=(4, 6)) - 1.0
    l0 = _pairwise_loss(mlp, X_fixed, X_accepted)
    for _ in range(300):
        mlp.train_step(X_fixed, X_accepted, lr=0.05, margin=MARGIN, reg_weight=REG_WEIGHT)
    l1 = _pairwise_loss(mlp, X_fixed, X_accepted)
    assert l1 < l0 * 0.5


def test_backward_matches_numerical_gradient():
    mlp = MLP(input_dim=4, hidden_dims=(3,), seed=7)
    rng = np.random.default_rng(0)
    X_fixed = rng.normal(size=(2, 4))
    X_accepted = rng.normal(size=(2, 4))
    eps = 1e-5
    lr = 1e-5
    grad = MLP(input_dim=4, hidden_dims=(3,), seed=7)
    grad.train_step(X_fixed, X_accepted, lr=lr, margin=MARGIN, reg_weight=REG_WEIGHT)
    for layer in range(len(mlp.W)):
        W0 = mlp.W[layer].copy()
        W1 = grad.W[layer]
        # train_step 做 W -= lr*dW（dW 是正向梯度）→ (W0-W1)/lr == +dL/dW，与数值梯度同号
        analytic = (W0 - W1) / lr
        num = np.zeros_like(W0)
        for i in range(W0.shape[0]):
            for j in range(W0.shape[1]):
                Wp = W0.copy()
                Wm = W0.copy()
                Wp[i, j] += eps
                Wm[i, j] -= eps
                mlp.W[layer] = Wp
                lp = _pairwise_loss(mlp, X_fixed, X_accepted)
                mlp.W[layer] = Wm
                lm = _pairwise_loss(mlp, X_fixed, X_accepted)
                num[i, j] = (lp - lm) / (2 * eps)
        mlp.W[layer] = W0
        np.testing.assert_allclose(analytic, num, rtol=1e-3, atol=1e-4)
        # bias 同样核对
        b0 = mlp.b[layer].copy()
        b1 = grad.b[layer]
        analytic_b = (b0 - b1) / lr
        num_b = np.zeros_like(b0)
        for k in range(b0.shape[0]):
            bp = b0.copy()
            bm = b0.copy()
            bp[k] += eps
            bm[k] -= eps
            mlp.b[layer] = bp
            lp = _pairwise_loss(mlp, X_fixed, X_accepted)
            mlp.b[layer] = bm
            lm = _pairwise_loss(mlp, X_fixed, X_accepted)
            num_b[k] = (lp - lm) / (2 * eps)
        mlp.b[layer] = b0
        np.testing.assert_allclose(analytic_b, num_b, rtol=1e-3, atol=1e-4)


def test_deterministic_training():
    rng = np.random.default_rng(1)
    Xf = rng.normal(size=(5, 8))
    Xa = rng.normal(size=(5, 8)) - 0.5
    a = MLP(input_dim=8, hidden_dims=(6,), seed=SEED)
    b = MLP(input_dim=8, hidden_dims=(6,), seed=SEED)
    for _ in range(50):
        a.train_step(Xf, Xa, lr=0.02, margin=MARGIN, reg_weight=REG_WEIGHT)
        b.train_step(Xf, Xa, lr=0.02, margin=MARGIN, reg_weight=REG_WEIGHT)
    for wa, wb in zip(a.W, b.W, strict=True):
        np.testing.assert_array_equal(wa, wb)
