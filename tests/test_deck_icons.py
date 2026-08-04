"""deck.render 图标后处理接线测试（monkeypatch 掉转换与真实后处理）。"""

import subprocess as _sp

import pytest

from offipy import charts, deck, icons


@pytest.fixture(autouse=True)
def _no_browser(monkeypatch):
    # 单测不真启动 chromium：render 的浏览器前置检查换成 no-op
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)


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

    def record(name):
        def f(*a, **k):
            calls.append(name)
            if name == "charts":
                assert a[0] == str(html) and a[1] == str(pptx)
            if name == "icons":
                assert a[0] == str(html) and a[1] == str(pptx)

        return f

    monkeypatch.setattr(charts, "postprocess_charts", record("charts"))
    monkeypatch.setattr(icons, "postprocess_icons", record("icons"))
    monkeypatch.setattr(
        deck.subprocess,
        "run",
        lambda *a, **k: _sp.CompletedProcess(args=[], returncode=0, stdout="", stderr=""),
    )

    out = deck.render(str(html), out=str(pptx), overwrite=True)  # placeholder 为后处理目标
    assert out == str(pptx)
    assert calls == ["charts", "icons"]


def test_render_charts_error_skips_icons(tmp_path, monkeypatch):
    """charts 后处理抛错时，icons 后处理不该被调用。"""
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
    monkeypatch.setattr(
        deck.subprocess,
        "run",
        lambda *a, **k: _sp.CompletedProcess(args=[], returncode=0, stdout="", stderr=""),
    )

    with pytest.raises(RuntimeError, match="charts failed"):
        deck.render(str(html), out=str(pptx), overwrite=True)  # placeholder 为后处理目标
    assert calls == ["charts"]  # icons 不该被调用
