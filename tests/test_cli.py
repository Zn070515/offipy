"""CLI 参数解析测试（不依赖 Office）。"""

from offipy.cli import build_parser


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
