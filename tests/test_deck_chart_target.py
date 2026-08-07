"""deck.render 图表后处理消费 target（注入副本）的回归测试。

Task 1 让 postprocess_charts 主题感知（读 HTML 的 `<style data-theme="NAME">`），
但那个 style 块只存在于 inject_theme/inject_layouts 产出的注入副本（target）里，
原始源 HTML 没有。本文件锁定：图表后处理拿到的是注入副本而非源文件，且临时副本用后即清。

monkeypatch 掉转换与真实后处理，单测不需 chromium，任何环境都可跑。
"""

import subprocess as _sp
from pathlib import Path

import pytest

from offipy import charts, deck


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


_SOURCE_HTML = (
    "<html><head></head><body><section data-pptx-slide><h2>x</h2></section></body></html>"
)


def _write_source(tmp_path):
    html = tmp_path / "d.html"
    html.write_text(_SOURCE_HTML, encoding="utf-8")
    pptx = tmp_path / "d.pptx"
    pptx.write_bytes(b"placeholder")
    return html, pptx


def test_theme_only_passes_injected_target_to_charts(tmp_path, monkeypatch):
    """theme= 注入时，图表后处理必须读注入副本（含 data-theme），而非源 HTML。"""
    html, pptx = _write_source(tmp_path)

    recorded: dict[str, object] = {}

    def record(*a, **k):
        recorded["html"] = a[0]
        recorded["exists_at_call"] = Path(a[0]).exists()
        recorded["content_at_call"] = Path(a[0]).read_text(encoding="utf-8")

    monkeypatch.setattr(charts, "postprocess_charts", record)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True, theme="mckinsey")
    assert out == str(pptx)

    recorded_path = recorded["html"]
    assert recorded_path != str(html)  # 是注入副本，不是源文件
    injected = Path(recorded_path)
    assert injected.name == "d.audited.html"  # 注入副本命名形状
    assert injected.parent.name.startswith("offipy-deck-")  # 临时目录前缀
    assert recorded["exists_at_call"] is True  # 后处理调用瞬间副本存在（读得到）
    assert 'data-theme="mckinsey"' in recorded["content_at_call"]  # 主题到达 charts
    # 注入副本随 TemporaryDirectory 用后即清
    assert not injected.exists()
    assert not injected.parent.exists()


def test_layouts_only_passes_injected_target_to_charts(tmp_path, monkeypatch):
    """apply_layouts= 注入时，图表后处理同样拿到临时注入副本并清理。"""
    html, pptx = _write_source(tmp_path)

    recorded: dict[str, object] = {}

    def record(*a, **k):
        recorded["html"] = a[0]
        recorded["exists_at_call"] = Path(a[0]).exists()

    monkeypatch.setattr(charts, "postprocess_charts", record)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True, apply_layouts=True)
    assert out == str(pptx)

    recorded_path = recorded["html"]
    assert recorded_path != str(html)  # 布局注入也产生临时副本
    injected = Path(recorded_path)
    assert injected.name == "d.audited.html"
    assert recorded["exists_at_call"] is True  # 后处理期间副本存在
    # 用后即清：render 返回后副本及其临时目录都不复存在
    assert not injected.exists()
    assert not injected.parent.exists()


def test_no_injection_passes_original_source_to_charts(tmp_path, monkeypatch):
    """无 theme/无 apply_layouts 时，target == html，源文件原样透传给图表后处理。"""
    html, pptx = _write_source(tmp_path)

    recorded: dict[str, object] = {}

    def record(*a, **k):
        recorded["html"] = a[0]

    monkeypatch.setattr(charts, "postprocess_charts", record)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True)
    assert out == str(pptx)

    assert recorded["html"] == str(html)  # 无注入路径行为不变


def test_theme_render_temp_target_exists_during_postprocess_and_cleaned(tmp_path, monkeypatch):
    """theme= 时，后处理调用瞬间注入副本必须存在，render 返回后整目录清理干净。"""
    html, pptx = _write_source(tmp_path)

    recorded: dict[str, object] = {}

    def record(*a, **k):
        recorded["html"] = a[0]
        recorded["exists_at_call"] = Path(a[0]).exists()
        recorded["parent_at_call"] = Path(a[0]).parent
        recorded["content_at_call"] = Path(a[0]).read_text(encoding="utf-8")

    monkeypatch.setattr(charts, "postprocess_charts", record)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True, theme="mckinsey")
    assert out == str(pptx)

    recorded_path = recorded["html"]
    assert recorded_path != str(html)
    assert recorded["exists_at_call"] is True  # 调用瞬间副本存在（后处理读得到）
    injected = Path(recorded_path)
    assert injected.parent.name.startswith("offipy-deck-")
    assert 'data-theme="mckinsey"' in recorded["content_at_call"]
    # render 返回后：注入副本与整个临时目录都已被清理
    assert not injected.exists()
    assert not injected.parent.exists()
    # 临时目录路径本身也消失了（不仅是文件被删）
    assert not recorded["parent_at_call"].exists()
