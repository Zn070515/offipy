"""deck.render 图标后处理接线测试（monkeypatch 掉转换与真实后处理）。"""

import subprocess as _sp
from pathlib import Path

import pytest

from offipy import charts, deck, icons


@pytest.fixture(autouse=True)
def _no_browser(monkeypatch):
    # 单测不真启动 chromium：render 的浏览器前置检查换成 no-op
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)


def _fake_run_creates_out(*a, **k):
    """mock subprocess.run：按 cmd 的 --out 产出临时 .pptx（原子替换要求 tmp 存在）。"""
    cmd = a[0]
    out = cmd[cmd.index("--out") + 1]
    Path(out).write_bytes(b"fake pptx")
    return _sp.CompletedProcess(args=[], returncode=0, stdout="", stderr="")


def test_render_wires_chart_then_icon_postprocess(tmp_path, monkeypatch):
    """render 依次调用 charts → icons 后处理（顺序 + 实参断言，不走真 convert/Playwright）。"""
    html = tmp_path / "d.html"
    html.write_text(
        "<html><body>"
        '<section class="slide" data-pptx-slide>'
        '<svg class="icon" data-icon="ph:check" viewBox="0 0 256 256"></svg>'
        "</section>"
        "</body></html>",
        encoding="utf-8",
    )
    pptx = tmp_path / "d.pptx"
    pptx.write_bytes(b"placeholder")

    calls: list[str] = []
    tmp_seen: dict[str, str] = {}

    def record(name):
        def f(*a, **k):
            calls.append(name)
            tmp_seen[name] = a[1]
            assert a[0] == str(html)
            # P0-6：后处理作用于临时 .pptx，不是最终路径
            assert Path(a[1]).name.startswith(".")
            assert Path(a[1]).name.endswith(".tmp.pptx")
            assert a[1] != str(pptx)

        return f

    monkeypatch.setattr(charts, "postprocess_charts", record("charts"))
    monkeypatch.setattr(icons, "postprocess_icons", record("icons"))
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True)  # placeholder 为后处理目标
    assert out == str(pptx)
    assert calls == ["charts", "icons"]
    assert tmp_seen["icons"] == tmp_seen["charts"]  # 两阶段作用于同一临时文件
    assert pptx.read_bytes() == b"fake pptx"  # 最终路径拿到转换产物
    assert not list(tmp_path.glob(".d.*.tmp.pptx"))  # 临时文件已清理


def test_render_charts_error_skips_icons(tmp_path, monkeypatch):
    """charts 后处理抛错时，icons 后处理不该被调用，且临时文件清理。"""
    html = tmp_path / "d.html"
    html.write_text("<section data-pptx-slide><h2>x</h2></section>", encoding="utf-8")
    pptx = tmp_path / "d.pptx"
    pptx.write_bytes(b"placeholder")

    calls: list[str] = []

    def boom(*a, **k):
        calls.append("charts")
        raise RuntimeError("charts failed")

    monkeypatch.setattr(charts, "postprocess_charts", boom)
    monkeypatch.setattr(icons, "postprocess_icons", lambda *a, **k: calls.append("icons"))
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_creates_out)

    with pytest.raises(RuntimeError, match="charts failed"):
        deck.render(str(html), out=str(pptx), overwrite=True)  # placeholder 为后处理目标
    assert calls == ["charts"]  # icons 不该被调用
    assert pptx.read_bytes() == b"placeholder"  # 失败不破坏已存在 .pptx
    assert not list(tmp_path.glob(".d.*.tmp.pptx"))  # 临时文件清理
