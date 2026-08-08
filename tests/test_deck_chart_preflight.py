"""deck.render chart-dominant 布局前置校验：未启用布局注入时 fail-fast。

Task 3：data-layout="chart-dominant" 的 .chart 可测尺寸由注入的布局 CSS 提供；
apply_layouts=False（CLI 缺 --layouts）时 CSS 未注入 → post-render 测量必然丢框 →
在启动 chromium / 跑 convert 之前就把用户拦下来，给可操作的修复指引。

monkeypatch 掉转换与真实后处理，单测不需 chromium，任何环境都可跑。
"""

import subprocess as _sp
from pathlib import Path

import pytest

from offipy import charts, deck
from offipy.exceptions import InvalidArgumentError


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


_CHART_DOM_HTML = (
    "<html><head></head><body>"
    '<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">'
    '<div class="chart-area"><div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div></div>'
    "</section>"
    "</body></html>"
)

_CUSTOM_CHART_HTML = (
    "<html><head></head><body>"
    '<section class="slide" data-pptx-slide>'
    '<div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div>'
    "</section>"
    "</body></html>"
)

_CHART_DOM_NO_CHART_HTML = (
    "<html><head></head><body>"
    '<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">'
    "<h2>title only</h2>"
    "</section>"
    "</body></html>"
)

_CHART_DOM_SINGLE_QUOTE_HTML = (
    "<html><head></head><body>"
    "<section class=\"slide chart-dominant\" data-pptx-slide data-layout='chart-dominant'>"
    '<div class="chart-area"><div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div></div>'
    "</section>"
    "</body></html>"
)

_TWO_CHART_DOM_HTML = (
    "<html><head></head><body>"
    '<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">'
    '<div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div>'
    "</section>"
    '<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">'
    '<div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div>'
    "</section>"
    "</body></html>"
)

_CHART_DOM_PLUS_PLAIN_HTML = (
    "<html><head></head><body>"
    '<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">'
    '<div class="chart" data-chart="bar" '
    'data-chart-data=\'{"categories":["a"],"series":[{"name":"s","values":[1]}]}\'></div>'
    "</section>"
    "<section data-pptx-slide><h2>plain</h2></section>"
    "</body></html>"
)

_CHART_DOM_SCRIPT_TARGET_HTML = (
    "<html><head></head><body>"
    '<section class="slide chart-dominant" data-pptx-slide data-layout="chart-dominant">'
    '<script type="application/json" data-chart-target="#c1">'
    '{"categories":["a"],"series":[{"name":"s","values":[1]}]}'
    "</script>"
    "</section>"
    "</body></html>"
)


def _write(html: Path, pptx: Path, text: str) -> None:
    html.write_text(text, encoding="utf-8")
    pptx.write_bytes(b"placeholder")


def test_chart_dominant_no_layouts_fails_before_converter(tmp_path, monkeypatch):
    """chart-dominant + apply_layouts=False → preflight 在 convert 子进程前抛错。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _CHART_DOM_HTML)

    run_calls: list[list] = []
    browser_calls: list[list] = []

    def recorder(*a, **k):
        run_calls.append(a)

    def browser_recorder(*a, **k):
        browser_calls.append(a)

    monkeypatch.setattr(deck, "_run_convert", recorder)
    monkeypatch.setattr(deck, "_preflight_browser", browser_recorder)

    with pytest.raises(InvalidArgumentError) as exc:
        deck.render(str(html), out=str(pptx), overwrite=True)
    msg = str(exc.value)
    assert "--layouts" in msg
    assert "[1]" in msg
    assert run_calls == []  # 转换子进程从未被调用（fail-fast）
    assert browser_calls == []  # 浏览器前置检查也没走到（preflight 更早拦下）


def test_chart_dominant_with_layouts_passes_preflight(tmp_path, monkeypatch):
    """chart-dominant + apply_layouts=True → preflight 放行，转换照常跑一次。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _CHART_DOM_HTML)

    run_calls: list[list] = []
    post_calls: list[list] = []

    def fake_run(*a, **k):
        run_calls.append(a)
        return _fake_run_creates_out(*a, **k)

    monkeypatch.setattr(deck, "_run_convert", fake_run)
    monkeypatch.setattr(charts, "postprocess_charts", lambda *a, **k: post_calls.append(a))

    out = deck.render(str(html), out=str(pptx), overwrite=True, apply_layouts=True)
    assert out == str(pptx)
    assert len(run_calls) == 1  # 转换子进程恰好一次
    assert len(post_calls) == 1  # 图表后处理也到达（走注入副本）


def test_custom_chart_no_layout_not_blocked(tmp_path, monkeypatch):
    """自定义布局的 .chart（未声明 chart-dominant）+ apply_layouts=False → 不被拦。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _CUSTOM_CHART_HTML)

    monkeypatch.setattr(charts, "postprocess_charts", lambda *a, **k: None)
    monkeypatch.setattr(deck, "_run_convert", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True)
    assert out == str(pptx)


def test_no_chart_no_behavior_change(tmp_path, monkeypatch):
    """无 .chart / 无 chart-dominant 的 deck → 与之前完全同路径，不被拦。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(
        html,
        pptx,
        "<html><head></head><body><section data-pptx-slide><h2>x</h2></section></body></html>",
    )

    monkeypatch.setattr(charts, "postprocess_charts", lambda *a, **k: None)
    monkeypatch.setattr(deck, "_run_convert", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True)
    assert out == str(pptx)


def test_chart_dominant_no_chart_inside_not_preflighted(tmp_path, monkeypatch):
    """声明 chart-dominant 但内部无图表 → helper 条件 (b) 不满足，preflight 不拦。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _CHART_DOM_NO_CHART_HTML)

    monkeypatch.setattr(charts, "postprocess_charts", lambda *a, **k: None)
    monkeypatch.setattr(deck, "_run_convert", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True)
    assert out == str(pptx)


def test_chart_dominant_single_quote_layout(tmp_path, monkeypatch):
    """单引号 data-layout='chart-dominant' 同样触发 preflight（两种引号都支持）。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _CHART_DOM_SINGLE_QUOTE_HTML)

    run_calls: list[list] = []
    monkeypatch.setattr(deck, "_run_convert", lambda *a, **k: run_calls.append(a))

    with pytest.raises(InvalidArgumentError) as exc:
        deck.render(str(html), out=str(pptx), overwrite=True)
    msg = str(exc.value)
    assert "--layouts" in msg
    assert "[1]" in msg
    assert run_calls == []  # fail-fast：convert 子进程未被调用


def test_chart_dominant_only_slides_filter(tmp_path, monkeypatch):
    """only_slides 把 preflight 收窄到实际渲染的页（未渲染的 chart-dominant 页被过滤）。"""
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _TWO_CHART_DOM_HTML)

    # only_slides=[1]：只渲染第 1 页 → 拦，且消息只列第 1 页（第 2 页被过滤）
    with pytest.raises(InvalidArgumentError) as exc:
        deck.render(str(html), out=str(pptx), overwrite=True, only_slides=[1])
    msg = str(exc.value)
    assert "[1]" in msg
    assert "[2]" not in msg

    # only_slides=[1,2]：两页都渲染 → 两页都列出
    with pytest.raises(InvalidArgumentError) as exc:
        deck.render(str(html), out=str(pptx), overwrite=True, only_slides=[1, 2])
    assert "[1, 2]" in str(exc.value)

    # 只渲染第 2 页（非 chart-dominant）→ 第 1 页被过滤 → preflight 放行
    mixed = tmp_path / "mixed.html"
    mixed_pptx = tmp_path / "mixed.pptx"
    _write(mixed, mixed_pptx, _CHART_DOM_PLUS_PLAIN_HTML)
    monkeypatch.setattr(charts, "postprocess_charts", lambda *a, **k: None)
    monkeypatch.setattr(deck, "_run_convert", _fake_run_creates_out)
    out = deck.render(str(mixed), out=str(mixed_pptx), overwrite=True, only_slides=[2])
    assert out == str(mixed_pptx)


def test_chart_dominant_script_target_not_preflighted(tmp_path, monkeypatch):
    """chart-dominant 页只有 data-chart-target 脚本、无真实图表容器 → 不拦。

    回归：条件 (b) 必须匹配真实图表容器（data-chart= / class="chart"），
    data-chart-target 里的子串 data-chart 不应误触发 preflight。
    """
    html = tmp_path / "d.html"
    pptx = tmp_path / "d.pptx"
    _write(html, pptx, _CHART_DOM_SCRIPT_TARGET_HTML)

    monkeypatch.setattr(charts, "postprocess_charts", lambda *a, **k: None)
    monkeypatch.setattr(deck, "_run_convert", _fake_run_creates_out)

    out = deck.render(str(html), out=str(pptx), overwrite=True)
    assert out == str(pptx)
