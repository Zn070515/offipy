"""deck 管线测试：theme 注入（不跑真实转换，拦截子进程）。"""

import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from offipy import deck


def _fake_run(created: dict):
    """拦截 convert 子进程：记录 cmd，按 convert 命名规则产出一个假 pptx。"""

    def fake_run(cmd, **kw):
        created["cmd"] = cmd
        created["injected_content"] = Path(cmd[2]).read_text(encoding="utf-8")
        inp = Path(cmd[2])  # cmd = [python, convert.py, html]
        out = None
        if "--out" in cmd:
            out = cmd[cmd.index("--out") + 1]
        elif inp.name.endswith(".audited.html"):
            out = str(inp.with_name(inp.name[: -len(".audited.html")] + ".pptx"))
        else:
            out = str(inp.with_suffix(".pptx"))
        Path(out).write_bytes(b"fake pptx")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    return fake_run


def test_render_no_theme_passes_original_html(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.render(str(html))
    assert created["cmd"][2] == str(html)
    assert pptx.endswith("deck.pptx")


def test_render_theme_injects_css_then_cleans_up(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><head><style data-theme="mckinsey"></style></head>'
        '<body><section class="slide" data-pptx-slide>hi</section></body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.render(str(html), theme="mckinsey")
    # 注入副本以 .audited.html 结尾（convert 跳过 work-copy），且已清理
    injected = created["cmd"][2]
    assert injected.endswith(".audited.html")
    assert not os.path.exists(injected), "注入副本应已清理"
    # 注入内容含主题 token（拦截时读取；render 返回后 tmp 已删）
    assert ":root {" in created["injected_content"]
    assert "--accent: #2251FF;" in created["injected_content"]
    # 输出名仍基于原 html
    assert pptx.endswith("deck.pptx")


def test_render_theme_unknown_raises(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><head></head><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    with pytest.raises(KeyError):
        deck.render(str(html), theme="nope")
    # 不触发任何转换
    assert "cmd" not in created


def test_make_passes_theme(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><head><style data-theme="dark-tech"></style></head><body>deck</body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    deck.make(str(html), open_live_flag=False, theme="dark-tech")
    assert "--accent: #38BDF8;" in created["injected_content"]  # dark-tech 强调色
