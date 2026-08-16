"""自定义 numpy MLP：配对 margin loss + centering 先验引导。

架构 [input_dim, 32, 16, 1]，ReLU 隐层，输出层线性（worth 标量）。纯手写
forward/backward（BP），无 autograd。配对监督：同一 (rule, profile) 组内
fixed 样本的 worth 应 > accepted 样本 + margin。centering 正则把固定样本 worth
拉向 +1、accepted 拉向 −1（先验：冷启动时学习路径输出 ≈ v2 手写行为）。

确定性：固定 SEED + 稳定遍历顺序 → 同一数据重复训练产出同一权重。golden
字节一致只在同一 numpy 版本 + 同一代码路径下成立（F2-C）。
"""

from __future__ import annotations

import math
from typing import Any

import numpy as np

SEED = 42
HIDDEN_DIMS = (32, 16)
MARGIN = 0.5
EPOCHS = 200
LR = 0.05
REG_WEIGHT = 0.2
# #112：全局梯度范数裁剪阈值。pair loss 无上界、reg 项随 worth 幅值放大梯度，
# 极端输入（NaN / 全零退化）会让梯度范数爆炸 → 裁剪到该上界防数值发散。
GRAD_CLIP_NORM = 5.0

# 形状各异的内部数组统一标注 float64；shape 维用 Any（mypy strict 下裸
# np.ndarray 触发 type-arg，与 vector.py 同一套泛型化约定）。
_Arr = np.ndarray[Any, np.dtype[np.float64]]

# #121 A3：容量自适应经验法则（Lunt & Xu 2016）的下限/上限与「加第二层」样本阈值。
H_MIN, H_MAX = 4, 32
SECOND_LAYER_N = 120
# #121 A3 / #134：samples_per_param 告警阈值——spp≥1 ok、≥0.25 warn、<0.25 critical。
# #134 重标定：旧 5/2 在全容量自适应（H=clamp(√n,4,32)）下误伤小样本——n=37 单层
# 25 params 曾得 0.34→critical，实际小样本 spp≥1 已是健康信号。建议样本数只作估算
# （suggest_n），容量仍是软告警不拒绝写盘。
_SPP_OK = 1.0
_SPP_WARN = 0.25


def adaptive_hidden_dims(n: int) -> tuple[int, ...]:
    """Lunt & Xu 2016 经验法则：H=clamp(round(√n),4,32)；n≥120 加第二层。"""
    H = int(max(H_MIN, min(H_MAX, round(math.sqrt(n)))))
    return (H, max(1, H // 2)) if n >= SECOND_LAYER_N else (H,)


def params_count(input_dim: int, hidden_dims: tuple[int, ...]) -> int:
    """逐层累加权重+偏置：input_dim→隐层→1 输出。"""
    dims = (input_dim, *hidden_dims, 1)
    return sum(dims[i] * dims[i + 1] + dims[i + 1] for i in range(len(dims) - 1))


def capacity_report(n: int, input_dim: int, hidden_dims: tuple[int, ...]) -> dict[str, Any]:
    """容量报告：samples_per_param 分级（ok/warn/critical）+ suggest_n 补样估算。"""
    params = params_count(input_dim, hidden_dims)
    spp = n / params if params else 0.0
    level = "ok" if spp >= _SPP_OK else ("warn" if spp >= _SPP_WARN else "critical")
    suggest_n = max(0, math.ceil(_SPP_OK * params) - n) if params else 0
    return {
        "n": n,
        "input_dim": input_dim,
        "hidden_dims": list(hidden_dims),
        "params": params,
        "samples_per_param": round(spp, 2),
        "level": level,
        "suggest_n": suggest_n,  # #134：达到 ok 约需再补的样本数（基于当前容量估算）
    }


class TrainingDiverged(RuntimeError):
    """训练数值发散（loss=inf/nan）——调用方据此返回 training_diverged 状态。"""


def train_mlp(X_fixed: _Arr, X_accepted: _Arr, hidden_dims: tuple[int, ...], *, seed: int) -> MLP:
    """按模块常量训练一个 member，返回训练好的 MLP；发散抛 TrainingDiverged。"""
    mlp = MLP(input_dim=X_fixed.shape[1], hidden_dims=hidden_dims, seed=seed)
    for _ in range(EPOCHS):
        loss = mlp.train_step(X_fixed, X_accepted, lr=LR, margin=MARGIN, reg_weight=REG_WEIGHT)
        if not math.isfinite(loss):
            raise TrainingDiverged()
    return mlp


def _clip_scale(dW: list[_Arr], db: list[_Arr]) -> float:
    """全局梯度范数裁剪系数：范数超 GRAD_CLIP_NORM 时缩放到上界，否则 1.0。"""
    grad_sq = sum(float((g**2).sum()) for g in dW) + sum(float((g**2).sum()) for g in db)
    grad_norm = math.sqrt(grad_sq) if grad_sq > 0.0 else 0.0
    if grad_norm > GRAD_CLIP_NORM:
        return GRAD_CLIP_NORM / grad_norm
    return 1.0


class MLP:
    def __init__(self, input_dim: int, hidden_dims: tuple[int, ...], seed: int = SEED) -> None:
        self.input_dim = input_dim
        self.hidden_dims = hidden_dims
        self.activations: list[_Arr] = []
        rng = np.random.default_rng(seed)
        dims = (input_dim, *hidden_dims, 1)
        self.W: list[_Arr] = []
        self.b: list[_Arr] = []
        for i in range(len(dims) - 1):
            fan_in, fan_out = dims[i], dims[i + 1]
            limit = np.sqrt(6.0 / (fan_in + fan_out))
            self.W.append(rng.uniform(-limit, limit, size=(fan_in, fan_out)))
            self.b.append(np.zeros(fan_out))

    def forward(self, X: _Arr) -> _Arr:
        """前向并缓存 activations（backward 用）。返回 worth (n,1)。"""
        self.activations = [X]
        a = X
        for i in range(len(self.W)):
            a = a @ self.W[i] + self.b[i]
            if i < len(self.W) - 1:
                a = np.maximum(a, 0.0)
            self.activations.append(a)
        return a

    def predict(self, x: _Arr) -> float:
        """单个特征向量 → worth 标量（不缓存 activations）。"""
        a = x.reshape(1, -1)
        for i in range(len(self.W)):
            a = a @ self.W[i] + self.b[i]
            if i < len(self.W) - 1:
                a = np.maximum(a, 0.0)
        return float(a[0, 0])

    def _predict_batch(self, X: _Arr) -> _Arr:
        a = X
        for i in range(len(self.W)):
            a = a @ self.W[i] + self.b[i]
            if i < len(self.W) - 1:
                a = np.maximum(a, 0.0)
        return a

    def predict_batch(self, X: _Arr) -> _Arr:
        """批量 forward 返回 worth (n,1)，不缓存 activations（推理/门禁用）。"""
        return self._predict_batch(X)

    def _backward(self, seed: _Arr) -> tuple[list[_Arr], list[_Arr]]:
        """反向传播：seed = dL/dA（输出层梯度，(n,1)）。返回 (dW, db)。"""
        a = self.activations
        d = seed
        dW: list[_Arr] = [np.zeros_like(w) for w in self.W]
        db: list[_Arr] = [np.zeros_like(v) for v in self.b]
        for i in range(len(self.W) - 1, -1, -1):
            dW[i] = a[i].T @ d
            db[i] = d.sum(axis=0)
            if i > 0:
                d = (d @ self.W[i].T) * (a[i] > 0)
        return dW, db

    def train_step(
        self,
        X_fixed: _Arr,
        X_accepted: _Arr,
        *,
        lr: float,
        margin: float,
        reg_weight: float,
    ) -> float:
        """一步梯度：配对 margin + centering 正则，返回 loss。"""
        n = X_fixed.shape[0]
        wf = self._predict_batch(X_fixed)
        wa = self._predict_batch(X_accepted)
        diff = wf - wa
        active = (diff < margin).astype(np.float64)
        pair_loss = float(np.maximum(margin - diff, 0.0).mean())
        reg_loss = float(((wf - 1.0) ** 2).mean() + ((wa + 1.0) ** 2).mean())
        loss = pair_loss + reg_weight * reg_loss
        seed_f = (-active + reg_weight * 2.0 * (wf - 1.0)) / n
        seed_a = (active + reg_weight * 2.0 * (wa + 1.0)) / n
        X = np.vstack([X_fixed, X_accepted])
        seed = np.vstack([seed_f, seed_a])
        self.forward(X)
        dW, db = self._backward(seed)
        scale = _clip_scale(dW, db)
        for i in range(len(self.W)):
            self.W[i] -= lr * scale * dW[i]
            self.b[i] -= lr * scale * db[i]
        return loss
