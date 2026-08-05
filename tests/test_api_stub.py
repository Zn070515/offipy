"""api.pyi 类型 stub 兜底测试：每个 schema op 都有 stub 方法、传输层参数、无漂移。"""

import importlib.util
import re
from pathlib import Path

from offipy import schema

ROOT = Path(__file__).resolve().parent.parent
STUB = ROOT / "src" / "offipy" / "api.pyi"
GEN = ROOT / "scripts" / "gen_api_stub.py"


def _gen():
    spec = importlib.util.spec_from_file_location("gen_api_stub", GEN)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _class_sigs(text: str, cls: str) -> dict[str, str]:
    """提取 class Xxx: 到下一个 class 之间的方法签名行：{op: 行文本}。"""
    m = re.search(rf"^class {cls}:.*?(?=^class |\Z)", text, re.MULTILINE | re.DOTALL)
    assert m, f"stub 缺少 class {cls}"
    sigs = {}
    for line in m.group(0).splitlines():
        mm = re.match(r"    def (\w+)\((.*)\) ->", line)
        if mm:
            sigs[mm.group(1)] = line.strip()
    return sigs


def test_every_schema_op_has_stub_method():
    gen = _gen()
    text = gen.render()
    for app in schema.apps():
        direct = {"excel": "Excel", "word": "Word", "ppt": "Ppt"}[app]
        remote = {"excel": "RemoteExcel", "word": "RemoteWord", "ppt": "RemotePpt"}[app]
        for op in schema.ops(app):
            assert op in _class_sigs(text, direct), f"{app}.{op} 缺 direct stub 方法"
            assert op in _class_sigs(text, remote), f"{app}.{op} 缺 remote stub 方法"


def test_transport_params_match_entrypoint():
    gen = _gen()
    text = gen.render()
    for app in schema.apps():
        direct = {"excel": "Excel", "word": "Word", "ppt": "Ppt"}[app]
        remote = {"excel": "RemoteExcel", "word": "RemoteWord", "ppt": "RemotePpt"}[app]
        for op in schema.ops(app):
            d_sig = _class_sigs(text, direct)[op]
            r_sig = _class_sigs(text, remote)[op]
            if op == "quit":
                continue  # quit 无 doc_id，destructive 元数据有但 stub 不加传输参数
            if schema.supports_expected_target(app, op):
                # 破坏性 op：remote 带 expected_target+follow_active，direct 只带 follow_active
                assert "follow_active: bool = False" in d_sig, f"{app}.{op} direct 缺 follow_active"
                assert "expected_target" not in d_sig, f"{app}.{op} direct 不应有 expected_target"
                assert "expected_target: dict | None = None" in r_sig
                assert "follow_active: bool = False" in r_sig
            else:
                assert "expected_target" not in d_sig
                assert "expected_target" not in r_sig


def test_stub_is_current_no_drift():
    gen = _gen()
    msg = "api.pyi 过期：跑 scripts/gen_api_stub.py 重新生成"
    assert STUB.read_text(encoding="utf-8") == gen.render(), msg
