"""deck.py 动画接入：参数穿透 / no_visual_audit 条件拒绝 / 告警码。"""

import json
from pathlib import Path

import pytest

from offipy.deck import _measure_warnings, _reject_no_visual_audit_declarations
from offipy.exceptions import InvalidArgumentError

MARKER_HTML = """<!doctype html><html><body>
<section data-pptx-slide>
  <div data-ppt-anim="fade">x</div>
</section></body></html>"""

TRANSITION_HTML = """<!doctype html><html><body>
<section data-pptx-slide data-ppt-transition="push">x</section></body></html>"""

FALLBACK_HTML = """<!doctype html><html><body>
<section data-pptx-slide><div data-aos="fade-up">x</div></section></body></html>"""

PLAIN_HTML = """<!doctype html><html><body>
<section data-pptx-slide><p>plain</p></section></body></html>"""


def test_reject_no_visual_audit_with_animations():
    with pytest.raises(InvalidArgumentError):
        _reject_no_visual_audit_declarations(MARKER_HTML, include_animations=True)


def test_reject_no_visual_audit_with_transition_marker():
    with pytest.raises(InvalidArgumentError):
        _reject_no_visual_audit_declarations(TRANSITION_HTML, include_animations=True)


def test_reject_no_visual_audit_with_fallback_marker():
    with pytest.raises(InvalidArgumentError):
        _reject_no_visual_audit_declarations(FALLBACK_HTML, include_animations=True)


def test_no_reject_without_animations_flag():
    _reject_no_visual_audit_declarations(MARKER_HTML)  # include_animations 默认 False


def test_no_reject_plain_html():
    _reject_no_visual_audit_declarations(PLAIN_HTML, include_animations=True)


def test_measure_warnings_anim_code():
    import pathlib
    import tempfile

    p = pathlib.Path(tempfile.mkdtemp()) / "measurements.json"
    p.write_text(json.dumps({"_warnings": [{"kind": "anim", "message": "skip"}]}), encoding="utf-8")
    ws = _measure_warnings(p)
    assert any(w.code == "deck.animation.skipped" for w in ws)


def test_render_tmp_animations_env_gate(monkeypatch, tmp_path):
    """animations=True → env 注入 OFFIPY_CONVERTER_ANIMATIONS；False → 无（逐字节不变）。"""
    import subprocess

    import offipy.deck as deck_mod

    html = tmp_path / "a.html"
    html.write_text(PLAIN_HTML, encoding="utf-8")
    captured = {}

    def fake_convert(cmd, timeout, env):
        captured["env"] = env
        tmp_pptx = cmd[cmd.index("--out") + 1]
        Path(tmp_pptx).write_bytes(b"pptx")
        return subprocess.CompletedProcess(cmd, 0, b"", b"")

    def fake_postprocess(label, fn, html, pptx):
        return {}

    monkeypatch.setattr(deck_mod, "_preflight_chart_layout", lambda *a: None)
    monkeypatch.setattr(deck_mod, "_preflight_browser", lambda: None)
    monkeypatch.setattr(deck_mod, "_run_convert", fake_convert)
    monkeypatch.setattr(deck_mod, "_postprocess", fake_postprocess)

    with deck_mod._render_tmp(
        str(html),
        None,
        None,
        False,
        600,
        None,
        False,
        False,
        animations=True,
        defer_audit_preserve=True,
    ):
        assert captured["env"].get("OFFIPY_CONVERTER_ANIMATIONS") == "1"

    captured.clear()
    with deck_mod._render_tmp(
        str(html),
        None,
        None,
        False,
        600,
        None,
        False,
        False,
        animations=False,
        defer_audit_preserve=True,
    ):
        assert "OFFIPY_CONVERTER_ANIMATIONS" not in captured["env"]


def test_render_tmp_animations_postprocess_only_when_enabled(monkeypatch, tmp_path):
    """animations=True → 后处理链含"动画"；False → 不含。"""
    import subprocess

    import offipy.deck as deck_mod

    html = tmp_path / "a.html"
    html.write_text(PLAIN_HTML, encoding="utf-8")
    labels = []

    def fake_convert(cmd, timeout, env):
        tmp_pptx = cmd[cmd.index("--out") + 1]
        Path(tmp_pptx).write_bytes(b"pptx")
        return subprocess.CompletedProcess(cmd, 0, b"", b"")

    def fake_postprocess(label, fn, html, pptx):
        labels.append(label)
        return {}

    monkeypatch.setattr(deck_mod, "_preflight_chart_layout", lambda *a: None)
    monkeypatch.setattr(deck_mod, "_preflight_browser", lambda: None)
    monkeypatch.setattr(deck_mod, "_run_convert", fake_convert)
    monkeypatch.setattr(deck_mod, "_postprocess", fake_postprocess)

    with deck_mod._render_tmp(
        str(html),
        None,
        None,
        False,
        600,
        None,
        False,
        False,
        animations=True,
        defer_audit_preserve=True,
    ):
        assert "动画" in labels

    labels.clear()
    with deck_mod._render_tmp(
        str(html),
        None,
        None,
        False,
        600,
        None,
        False,
        False,
        animations=False,
        defer_audit_preserve=True,
    ):
        assert "动画" not in labels
