"""vendored 转换器修复的单测（0.10.2）：work-copy mtime 重建 + 破图 fail-fast。

不触 chromium：直接 importlib 加载 convert.py 纯函数（_work_copy_target /
_collect_broken_images）验证决策逻辑，不跑完整转换流水线。
"""

import importlib.util
import sys
import time
from pathlib import Path

import pytest

_VENDOR = (
    Path(__file__).resolve().parent.parent / "src" / "offipy" / "_vendor" / "html_to_editable_pptx"
)
_SCRIPTS = _VENDOR / "scripts"


@pytest.fixture(scope="module")
def convert_mod():
    # convert.py 顶层 import SCRIPTS 下的 font_paths/text_utils，需先注入 sys.path
    # （convert.py 自己也会 insert，但 spec 加载前 import 解析需可见）。
    if str(_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(_SCRIPTS))
    spec = importlib.util.spec_from_file_location(
        "offipy_convert_under_test", _VENDOR / "convert.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_work_copy_creates_then_reuses_then_rebuilds(convert_mod, tmp_path):
    src = tmp_path / "deck.html"
    src.write_text("v1", encoding="utf-8")
    audited = tmp_path / "deck.audited.html"

    # 首次：创建副本
    assert convert_mod._work_copy_target(src) == audited
    assert audited.read_text(encoding="utf-8") == "v1"

    # 源未变：复用旧副本（保留手动 audit 修复）
    audited.write_text("manual fix", encoding="utf-8")
    assert convert_mod._work_copy_target(src) == audited
    assert audited.read_text(encoding="utf-8") == "manual fix"

    # 源更新（mtime 更大）：重建副本，改源即刻生效，不静默复用旧副本
    time.sleep(0.02)  # 保证 mtime 分辨率内源严格更新
    src.write_text("v2", encoding="utf-8")
    assert convert_mod._work_copy_target(src) == audited
    assert audited.read_text(encoding="utf-8") == "v2"


def test_work_copy_non_html_passthrough(convert_mod, tmp_path):
    other = tmp_path / "notes.txt"
    other.write_text("x", encoding="utf-8")
    assert convert_mod._work_copy_target(other) == other
    assert not (tmp_path / "notes.audited.txt").exists()


def test_work_copy_audited_input_passthrough(convert_mod, tmp_path):
    audited = tmp_path / "deck.audited.html"
    audited.write_text("fix", encoding="utf-8")
    # 显式传 .audited.html → 不切换、不重建
    assert convert_mod._work_copy_target(audited) == audited
    assert audited.read_text(encoding="utf-8") == "fix"


def test_collect_broken_images_slides_form(convert_mod):
    meas = {
        "slides": [
            {
                "records": [
                    {"kind": "img", "src": "fig/ok.png", "imgBroken": False},
                    {"kind": "img", "src": "fig/missing.png", "imgBroken": True},
                    {"kind": "text", "text": "hi"},
                    {"kind": "img", "src": "http://x/y.png", "imgBroken": True},
                ]
            },
            {"records": []},
        ]
    }
    broken = convert_mod._collect_broken_images(meas)
    assert [b["src"] for b in broken] == ["fig/missing.png", "http://x/y.png"]


def test_collect_broken_images_single_page_form(convert_mod):
    # 单页 measure 返回 {slide, records}（兼容旧 API）
    meas = {"slide": 1, "records": [{"kind": "img", "src": "a.png", "imgBroken": True}]}
    assert [b["src"] for b in convert_mod._collect_broken_images(meas)] == ["a.png"]


def test_collect_broken_images_empty(convert_mod):
    assert convert_mod._collect_broken_images({"slides": []}) == []
    assert convert_mod._collect_broken_images({}) == []
