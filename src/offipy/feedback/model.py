"""model.json 读写 + schema 版本门禁。

- 原子写：临时文件 + os.replace（进程中断不留半截文件）
- F2-E：失败路径（调用方异常 / _fail 注入）不触碰旧文件，只有成功才 replace
- 门禁：模型可用性只由 input_schema_version 决定；output_schema_version 仅记录
- 损坏 / 缺失 → load_model 返回 None → 冷启动回退 v2
"""

from __future__ import annotations

import contextlib
import json
import os
import tempfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .mlp import MLP

MODEL_FILE = "art_feedback_model.json"
MODEL_FORMAT_VERSION = 2

REQUIRED_KEYS = (
    "schema_version",
    "input_schema_version",
    "output_schema_version",
    "members",
    "preprocessing",
    "calibration",
    "abstain",
    "stats",
)


def model_file(feedback_dir: str | Path | None = None) -> Path:
    d = Path(feedback_dir) if feedback_dir else Path.home() / ".offipy"
    return d / MODEL_FILE


def _weights_to_dict(mlp: MLP) -> dict[str, list[Any]]:
    return {
        "W": [w.tolist() for w in mlp.W],
        "b": [b.tolist() for b in mlp.b],
    }


def save_model(
    members: list[tuple[int, MLP]],
    *,
    input_schema_version: str,
    output_schema_version: str,
    seed: int,
    stats: dict[str, Any],
    preprocessing: dict[str, Any],
    calibration: dict[str, Any],
    abstain: dict[str, Any],
    path: Path,
    _fail: bool = False,
) -> Path:
    """把 ensemble members + 预处理/校准元数据原子写到 path。返回 path。"""
    data = {
        "schema_version": MODEL_FORMAT_VERSION,
        "input_schema_version": input_schema_version,
        "output_schema_version": output_schema_version,
        "seed": seed,
        "members": [
            {"seed": s, "hidden_dims": list(m.hidden_dims), "weights": _weights_to_dict(m)}
            for s, m in members
        ],
        "preprocessing": preprocessing,
        "calibration": calibration,
        "abstain": abstain,
        "trained_at": _now_iso(),
        "stats": stats,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    if _fail:
        raise OSError("injected save failure (F2-E test hook)")
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".model-", suffix=".tmp")
    tmp_path = Path(tmp)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(payload)
        tmp_path.replace(path)
    except BaseException:
        with contextlib.suppress(OSError):
            tmp_path.unlink()
        raise
    return path


def _now_iso() -> str:
    from datetime import datetime, timezone

    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def load_model(path: Path) -> dict[str, Any] | None:
    """读 model.json；缺失 / 损坏 / 缺关键字段 → None。"""
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, UnicodeError, OSError):
        return None
    if not isinstance(data, dict):
        return None
    if not all(k in data for k in REQUIRED_KEYS):
        return None
    return data


def model_valid(data: dict[str, Any], input_schema_version: str) -> bool:
    """模型可用性：input_schema_version 匹配才有效（权重维度才对齐）。"""
    return data.get("input_schema_version") == input_schema_version


def kept_valid(pre: dict[str, Any], num_features: int) -> bool:
    """kept 是否为当前 feature_keys 下的合法绝对下标列表。

    #150：schema bump 后 persisted kept 可能越界 / 缺失 / 含非数值 / 负值
    → 视为无模型回退 v2。全 int（不含 bool）、非空、0 ≤ i < num_features。
    """
    kept = pre.get("kept")
    if not isinstance(kept, list) or not kept:
        return False
    return all(
        isinstance(i, int) and not isinstance(i, bool) and 0 <= i < num_features for i in kept
    )


def weights_from_dict(data: dict[str, Any], *, input_dim: int, hidden_dims: tuple[int, ...]) -> MLP:
    """按 data['weights'] 还原 MLP（形状校验失败抛 ValueError → 调用方视为无模型）。"""
    import numpy as np

    from .mlp import MLP

    if data.get("hidden_dims") != list(hidden_dims):
        raise ValueError(f"hidden_dims 不匹配: {data.get('hidden_dims')} vs {list(hidden_dims)}")
    mlp = MLP(input_dim=input_dim, hidden_dims=hidden_dims, seed=int(data["seed"]))
    weights = data["weights"]
    Ws = [np.asarray(w, dtype=np.float64) for w in weights["W"]]
    bs = [np.asarray(v, dtype=np.float64) for v in weights["b"]]
    for i, (W, b) in enumerate(zip(Ws, bs, strict=True)):
        if W.shape != mlp.W[i].shape or b.shape != mlp.b[i].shape:
            raise ValueError(f"权重形状不匹配 layer {i}: {W.shape} vs {mlp.W[i].shape}")
        mlp.W[i] = W
        mlp.b[i] = b
    return mlp
