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


def test_deck_outline_writes_html(tmp_path):
    from offipy import cli

    md = tmp_path / "outline.md"
    md.write_text("# T\n> S\n\n## 页一 @layout: big-number\n- 甲\n", encoding="utf-8")
    out = tmp_path / "deck.html"
    cli.main(["deck", "outline", "--input", str(md), "--theme", "mckinsey", "--out", str(out)])
    html = out.read_text(encoding="utf-8")
    assert "data-pptx-slide" in html
    assert 'data-layout="big-number"' in html
    assert '<style data-theme="mckinsey">' in html


def test_deck_outline_prints_json_without_out(tmp_path, capsys):
    from offipy import cli

    md = tmp_path / "outline.md"
    md.write_text("# T\n\n## 页\n- 甲\n", encoding="utf-8")
    cli.main(["deck", "outline", "--input", str(md)])
    assert '"title": "T"' in capsys.readouterr().out
