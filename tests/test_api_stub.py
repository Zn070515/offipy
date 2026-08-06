"""api.pyi 类型 stub 兜底测试：每个 schema op 都有 stub 方法、传输层参数、无漂移。"""

import importlib.util
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

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


# --- P1-6：read_slide_texts 破坏性重构的类型门禁 ---


def test_read_slide_texts_stub_signature_snapshot():
    # 返回 list[SlideTextRecord] 而非 list[dict]；keyword-only 参数忠实保留
    gen = _gen()
    text = gen.render()
    direct = _class_sigs(text, "Ppt")["read_slide_texts"]
    assert "-> list[SlideTextRecord]" in direct
    assert "list[dict]" not in direct
    assert direct.startswith(
        "def read_slide_texts(self, slide_idx: int, *, include_empty: bool = False"
    ), direct
    assert ", recursive: bool = True" in direct
    # #25：只读 op 也带 follow_active（read_slide_texts 放行）
    assert ", doc_id: str | None = None, follow_active: bool = False)" in direct
    # 旧摘要形态（无 slide_idx）已移除，摘要语义移到 read_slide_summary
    assert "read_slide_summary" in _class_sigs(text, "Ppt")
    summary = _class_sigs(text, "Ppt")["read_slide_summary"]
    assert (
        "def read_slide_summary(self, doc_id: str | None = None, *, follow_active: bool = False)"
        in summary
    )
    # 远程 facade 同样带 keyword-only request_id
    remote = _class_sigs(text, "RemotePpt")["read_slide_texts"]
    assert "request_id: str | None = None" in remote


def test_mypy_user_example_reveals_sliderecord_types(tmp_path):
    # P1-6：真实用户代码在 api.pyi 上的类型推导门禁。SlideTextRecord 是 TypedDict，
    # 字段类型（int/str/Literal）必须对 mypy 可见——退化成 list[dict] 就是 Any。
    snippet = (
        "from offipy import Ppt\n\n"
        "records = Ppt().read_slide_texts(1)\n"
        "reveal_type(records)\n"
        "reveal_type(records[0]['shape_id'])\n"
        "reveal_type(records[0]['text'])\n"
        "reveal_type(records[0]['placeholder_type'])\n"
        "reveal_type(records[0]['coordinate_space'])\n"
    )
    src_file = tmp_path / "user_example.py"
    src_file.write_text(snippet, encoding="utf-8")

    env = os.environ.copy()
    env["MYPYPATH"] = str(ROOT / "src")
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = _mypy_cmd() + [
        "--platform",
        "win32",
        "--ignore-missing-imports",
        "--no-incremental",
        str(src_file),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=180)
    except FileNotFoundError:
        pytest.skip("环境无 mypy")
    out = proc.stdout + proc.stderr
    if "No module named mypy" in out:
        pytest.skip("环境无 mypy")  # python -m mypy 会正常启动，只是解释器找不到模块
    if "Cannot find implementation or library stub" in out:
        pytest.skip(f"mypy 无法解析 offipy 类型（{out.strip()}）")
    assert "SlideTextRecord" in out, out
    assert 'Revealed type is "int"' in out, out
    assert 'Revealed type is "str"' in out, out
    assert "Literal['slide']" in out, out  # coordinate_space 类型流透传，非 str/Any
    assert 'Revealed type is "Any"' not in out, out


def _mypy_cmd() -> list[str]:
    exe = shutil.which("mypy")
    if exe:
        return [exe]
    return [sys.executable, "-m", "mypy"]
