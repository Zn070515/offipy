"""docs/api 英文翻译层覆盖断言：schema 每个 op 都有英文描述，且无孤儿条目。

schema.py 的描述保持中文单源（MCP/CLI 工具描述不变）；英文版由
`scripts/gen_api_ref.py` 的 `_EN_DESC` 提供。新增 op 忘补英文描述时，
本测试直接失败（生成器 main() 的 `_guard_en_coverage` 也会在生成时失败）。
"""

import importlib.util
from pathlib import Path

from offipy import schema

_SCRIPT = Path(__file__).resolve().parent.parent / "scripts" / "gen_api_ref.py"


def _load_gen():
    spec = importlib.util.spec_from_file_location("gen_api_ref_en_test", _SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod


def test_every_schema_op_has_english_description():
    gen = _load_gen()
    for app in schema.apps():
        for op in schema.ops(app):
            assert (app, op) in gen._EN_DESC, f"{app}.{op} 缺英文描述（_EN_DESC）"


def test_no_orphan_english_descriptions():
    gen = _load_gen()
    for key in gen._EN_DESC:
        app, op = key
        assert app in schema.apps(), f"孤儿英文条目 {key}: 未知应用"
        assert op in schema.ops(app), f"孤儿英文条目 {key}: schema 无此 op"


def test_en_coverage_guard_passes():
    # _guard_en_coverage 对当前 schema 应无缺漏（与 test_every_* 一致但不 import 冲突）
    gen = _load_gen()
    missing = [
        f"{a}.{o}" for a in schema.apps() for o in schema.ops(a) if (a, o) not in gen._EN_DESC
    ]
    assert missing == []
