"""deck 资产管线集成：目标注入 / no_visual_audit 前置 / only_slides 一致性。

A3 Task 3：asset 声明（data-asset/data-primitive/legacy data-icon）与图表一样
需要注入副本 target（converter 量到注入后的确定性 ID），且 no_visual_audit 的
前置检查要在 chromium / convert 子进程之前拦下非法组合。单测 monkeypatch 浏览器
与转换子进程，任何环境可跑。
"""

import subprocess as _sp
from pathlib import Path

import pytest

from offipy import charts, deck


@pytest.fixture(autouse=True)
def _no_browser(monkeypatch):
    # 单测不真启动 chromium：render 的浏览器前置检查换成 no-op（重测 preflight 顺序时覆盖）
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)


def _fake_run(recorded: dict):
    """mock subprocess.run：按 cmd 的 --out 产出临时 .pptx，记录 cmd。"""

    def fake_run(cmd, **kw):
        recorded["cmd"] = cmd
        out = cmd[cmd.index("--out") + 1]
        Path(out).write_bytes(b"fake pptx")
        return _sp.CompletedProcess(args=[], returncode=0, stdout="", stderr="")

    return fake_run


def _record_target(recorded: dict):
    """mock postprocess_charts：记录收到的 html target 路径 + 调用瞬间内容。

    注入副本在 render 返回时已被 finally 清理，内容必须在 postprocess 调用瞬间读。
    """

    def record(html_path, pptx_path):
        recorded["html"] = html_path
        recorded["exists_at_call"] = Path(html_path).exists()
        recorded["content_at_call"] = Path(html_path).read_text(encoding="utf-8")

    return record


_HTML = """<!doctype html>
<html><head></head><body>
<section data-pptx-slide><h1>一</h1></section>
<section data-pptx-slide><div data-asset="asset://ph/icon/check"></div></section>
</body></html>
"""


def _write(tmp_path, text=_HTML):
    html = tmp_path / "d.html"
    html.write_text(text, encoding="utf-8")
    pptx = tmp_path / "d.pptx"
    pptx.write_bytes(b"placeholder")
    return html, pptx


# ---------------------------------------------------------------------------
# no_visual_audit 前置：声明 → chromium / convert 之前 fail-fast
# ---------------------------------------------------------------------------


def _assert_nova_rejected_before_browser(tmp_path, monkeypatch, html_text, marker):
    html, pptx = _write(tmp_path, html_text)
    browser_calls: list = []
    run_calls: list = []

    monkeypatch.setattr(deck, "_preflight_browser", lambda *a, **k: browser_calls.append(a))
    monkeypatch.setattr(deck.subprocess, "run", lambda *a, **k: run_calls.append(a))

    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(html), out=str(pptx), overwrite=True, no_visual_audit=True)
    assert marker in str(exc.value)  # 信息列出具体声明类型
    assert browser_calls == []  # chromium 从未启动
    assert run_calls == []  # convert 子进程从未调用


def test_no_visual_audit_rejects_data_asset_before_browser(tmp_path, monkeypatch):
    _assert_nova_rejected_before_browser(tmp_path, monkeypatch, _HTML, "data-asset")


def test_no_visual_audit_rejects_data_primitive_before_browser(tmp_path, monkeypatch):
    _assert_nova_rejected_before_browser(
        tmp_path,
        monkeypatch,
        '<section data-pptx-slide><div data-primitive="quote-mark"></div></section>',
        "data-primitive",
    )


def test_no_visual_audit_still_rejects_legacy_data_icon_before_browser(tmp_path, monkeypatch):
    _assert_nova_rejected_before_browser(
        tmp_path,
        monkeypatch,
        '<section data-pptx-slide><svg data-icon="ph:check"></svg></section>',
        "data-icon",
    )


def test_no_visual_audit_still_rejects_data_chart_before_browser(tmp_path, monkeypatch):
    _assert_nova_rejected_before_browser(
        tmp_path,
        monkeypatch,
        '<section data-pptx-slide><div data-chart="bar"></div></section>',
        "data-chart",
    )


def test_no_visual_audit_message_lists_declaration_type(tmp_path, monkeypatch):
    """data-asset 单独出现时，信息点名的就是它（不混入无关类型）。"""
    html, pptx = _write(tmp_path)
    monkeypatch.setattr(deck.subprocess, "run", lambda *a, **k: None)

    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(html), out=str(pptx), overwrite=True, no_visual_audit=True)
    msg = str(exc.value)
    assert "data-asset" in msg
    assert "data-chart" not in msg  # 未声明图表就不误报图表
    assert "data-icon" not in msg
    assert "data-primitive" not in msg


def test_no_visual_audit_plain_deck_still_ok(tmp_path, monkeypatch):
    """回归：无任何声明的纯 deck 配 no_visual_audit 不受前置校验影响。"""
    html, pptx = _write(
        tmp_path, "<html><body><section data-pptx-slide><p>ok</p></section></body></html>"
    )
    recorded = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(recorded))

    out = deck.render(str(html), out=str(pptx), overwrite=True, no_visual_audit=True)
    assert out == str(pptx)
    assert "--no-visual-audit" in recorded["cmd"]  # 转换器确实收到该开关


# ---------------------------------------------------------------------------
# 注入副本 target：asset 声明强制临时 target，无 theme/layout 也如此
# ---------------------------------------------------------------------------


def test_asset_only_passes_temp_injected_target(tmp_path, monkeypatch):
    """无 theme/layout 时，asset 声明也强制临时注入副本；源文件不动、用后即清。"""
    html, pptx = _write(tmp_path)
    original = html.read_text(encoding="utf-8")

    recorded = {}
    monkeypatch.setattr(charts, "postprocess_charts", _record_target(recorded))
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(recorded))

    out = deck.render(str(html), out=str(pptx), overwrite=True)
    assert out == str(pptx)

    target = Path(recorded["html"])
    assert str(target) != str(html)  # 注入副本而非源文件
    assert target.name == "d.audited.html"  # 命名让 convert 跳过 work-copy 分支
    assert target.parent.name.startswith("offipy-deck-")  # 临时目录前缀
    assert recorded["exists_at_call"] is True  # 后处理瞬间副本存在
    assert 'data-offipy-asset-id="asset-s02-001"' in recorded["content_at_call"]
    # 用户源 HTML 未被就地编辑
    assert html.read_text(encoding="utf-8") == original
    # 注入副本随 TemporaryDirectory 用后即清
    assert not target.exists()
    assert not target.parent.exists()


def test_asset_only_no_visual_audit_rejected_before_temp_target(tmp_path, monkeypatch):
    """no_visual_audit + asset 声明在临时 target 生成前就拦下（无残留副本）。"""
    html, pptx = _write(tmp_path)
    monkeypatch.setattr(deck.subprocess, "run", lambda *a, **k: None)

    with pytest.raises(deck.InvalidArgumentError):
        deck.render(str(html), out=str(pptx), overwrite=True, no_visual_audit=True)
    assert not list(tmp_path.glob("offipy-deck-*"))


def test_theme_plus_asset_target_has_theme_and_ids(tmp_path, monkeypatch):
    """theme + asset：注入副本同时带主题 token 与确定性 ID（顺序：先注入后解析）。"""
    html, pptx = _write(tmp_path)
    recorded = {}
    monkeypatch.setattr(charts, "postprocess_charts", _record_target(recorded))
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(recorded))

    out = deck.render(str(html), out=str(pptx), overwrite=True, theme="mckinsey")
    assert out == str(pptx)

    content = recorded["content_at_call"]
    assert 'data-theme="mckinsey"' in content  # 主题注入在 asset 解析前完成
    assert 'data-offipy-asset-id="asset-s02-001"' in content


def test_asset_target_cleaned_on_failure(tmp_path, monkeypatch):
    """后处理抛错时，注入副本随 finally 清理，不留残留。"""
    html, pptx = _write(tmp_path)
    target_seen: dict[str, str] = {}

    def boom(html_path, pptx_path):
        target_seen["html"] = html_path
        raise RuntimeError("boom")

    monkeypatch.setattr(charts, "postprocess_charts", boom)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run({}))

    with pytest.raises(deck.ConversionError):
        deck.render(str(html), out=str(pptx), overwrite=True)
    assert not Path(target_seen["html"]).exists()
    assert not Path(target_seen["html"]).parent.exists()


# ---------------------------------------------------------------------------
# only_slides：声明页码与 converter 一致
# ---------------------------------------------------------------------------


def test_only_slides_keeps_full_html_declaration_ordinals(tmp_path, monkeypatch):
    """only_slides 只重渲子集，但声明 ID 仍按全 HTML 页码编号（converter 对齐）。"""
    html, pptx = _write(tmp_path)
    recorded = {}
    monkeypatch.setattr(charts, "postprocess_charts", _record_target(recorded))
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(recorded))

    out = deck.render(str(html), out=str(pptx), overwrite=True, only_slides=[2])
    assert out == str(pptx)

    cmd = recorded["cmd"]
    assert cmd[cmd.index("--only-slides") + 1] == "2"
    content = recorded["content_at_call"]
    # 声明按全 HTML 第 2 页编号（asset-s02-001），不被 only_slides 重排
    assert 'data-offipy-asset-id="asset-s02-001"' in content
