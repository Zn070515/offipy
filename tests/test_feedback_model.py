"""model.json 原子读写 + input_schema_version 门禁 + F2-E 保留旧模型。"""

import numpy as np
import pytest

from offipy.art.features_registry import feature_keys, feature_schema_version
from offipy.feedback.mlp import HIDDEN_DIMS, MLP
from offipy.feedback.model import (
    MODEL_FILE,
    MODEL_FORMAT_VERSION,
    load_model,
    model_file,
    model_valid,
    save_model,
    weights_from_dict,
)


def _mlp():
    return MLP(input_dim=len(feature_keys()), hidden_dims=HIDDEN_DIMS, seed=42)


def test_model_file_default_location(tmp_path):
    assert model_file(tmp_path) == tmp_path / MODEL_FILE


def test_save_load_roundtrip(tmp_path):
    mlp = _mlp()
    path = save_model(
        mlp,
        input_schema_version=feature_schema_version(),
        output_schema_version="1",
        seed=42,
        hidden_dims=HIDDEN_DIMS,
        stats={"pairs": 10, "samples": 20, "loss": 0.1, "rules_with_pairs": 2},
        path=model_file(tmp_path),
    )
    data = load_model(path)
    assert data is not None
    assert data["schema_version"] == MODEL_FORMAT_VERSION
    assert data["input_schema_version"] == feature_schema_version()
    assert data["output_schema_version"] == "1"
    assert data["seed"] == 42
    assert data["hidden_dims"] == list(HIDDEN_DIMS)
    assert data["stats"]["pairs"] == 10
    restored = weights_from_dict(data, input_dim=len(feature_keys()), hidden_dims=HIDDEN_DIMS)
    for w0, w1 in zip(mlp.W, restored.W, strict=True):
        np.testing.assert_array_equal(w0, w1)
    for b0, b1 in zip(mlp.b, restored.b, strict=True):
        np.testing.assert_array_equal(b0, b1)


def test_save_is_atomic_no_partial_file(tmp_path):
    path = model_file(tmp_path)
    save_model(
        _mlp(),
        input_schema_version="1",
        output_schema_version="1",
        seed=1,
        hidden_dims=HIDDEN_DIMS,
        stats={},
        path=path,
    )
    # 目录里只有最终文件（临时文件已 replace，无残留）
    assert [p.name for p in tmp_path.iterdir()] == [MODEL_FILE]


def test_load_corrupt_returns_none(tmp_path):
    path = model_file(tmp_path)
    path.write_text("not-json{", encoding="utf-8")
    assert load_model(path) is None
    path.write_text('{"weights": [1]}', encoding="utf-8")  # 缺关键字段
    assert load_model(path) is None


def test_load_missing_returns_none(tmp_path):
    assert load_model(model_file(tmp_path)) is None


def test_model_valid_gates_on_input_schema_version(tmp_path):
    path = save_model(
        _mlp(),
        input_schema_version=feature_schema_version(),
        output_schema_version="1",
        seed=42,
        hidden_dims=HIDDEN_DIMS,
        stats={},
        path=model_file(tmp_path),
    )
    data = load_model(path)
    assert data is not None
    assert model_valid(data, feature_schema_version()) is True
    assert model_valid(data, "999") is False  # 过期


def test_save_failure_keeps_old_model(tmp_path):
    """F2-E：失败路径不删除、不覆盖旧 model.json。"""
    old = save_model(
        _mlp(),
        input_schema_version=feature_schema_version(),
        output_schema_version="1",
        seed=42,
        hidden_dims=HIDDEN_DIMS,
        stats={},
        path=model_file(tmp_path),
    )
    before = old.read_bytes()
    with old.open("ab"):  # touch，确认存在
        pass
    # 模拟失败：_fail 注入的 OSError 在 mkstemp/replace 之前抛出，旧模型不被触碰
    with pytest.raises(OSError):
        save_model(
            _mlp(),
            input_schema_version=feature_schema_version(),
            output_schema_version="1",
            seed=1,
            hidden_dims=HIDDEN_DIMS,
            stats={},
            path=old,
            _fail=True,
        )
    assert old.read_bytes() == before  # 旧模型未被覆盖
