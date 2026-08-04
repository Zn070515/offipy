"""CLI 参数解析测试（不依赖 Office）。"""

import pytest

from offipy.cli import build_parser


def test_no_subcommand_errors_with_usage():
    # 回归：无子命令时曾抛 AttributeError（args 无 kwargs 属性），
    # 现在 argparse 报 usage 并退出。
    with pytest.raises(SystemExit):
        build_parser().parse_args([])


def test_quit_target_does_not_shadow_app():
    # 回归：quit 子命令的参数名曾与顶层 dest "app" 冲突，
    # 导致 args.app 被覆盖成 "ppt" 而走错分支。
    args = build_parser().parse_args(["quit", "ppt"])
    assert args.app == "quit"
    assert args.target == "ppt"


def test_office_subcommand_has_kwargs():
    args = build_parser().parse_args(["ppt", "add_slide", "--layout", "2"])
    assert args.app == "ppt"
    assert args.kwargs == ["--layout", "2"]


def test_parse_kwargs_flag_without_value():
    from offipy.cli import _parse_kwargs

    assert _parse_kwargs(["--no-open"]) == {"no-open": True}
    assert _parse_kwargs(["--no-open", "--out", "x.pptx"]) == {
        "no-open": True,
        "out": "x.pptx",
    }
    assert _parse_kwargs(["--html", "a.html"]) == {"html": "a.html"}


def test_deck_make_passes_theme_and_layouts(monkeypatch, capsys):
    from offipy import cli

    captured = {}

    def fake_make(html, **kw):
        captured["html"] = html
        captured["kw"] = kw
        return r"C:\out\deck.pptx"

    monkeypatch.setattr("offipy.deck.make", fake_make)
    cli.main(["deck", "make", "--html", "x.html", "--theme", "mckinsey", "--layouts"])
    assert captured["html"] == "x.html"
    assert captured["kw"]["theme"] == "mckinsey"
    assert captured["kw"]["apply_layouts"] is True
    assert "deck.pptx" in capsys.readouterr().out
