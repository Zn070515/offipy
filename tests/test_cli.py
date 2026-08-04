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


def test_deck_make_passes_overwrite(monkeypatch, capsys):
    from offipy import cli

    captured = {}

    def fake_make(html, **kw):
        captured["kw"] = kw
        return r"C:\out\deck.pptx"

    monkeypatch.setattr("offipy.deck.make", fake_make)
    cli.main(["deck", "make", "--html", "x.html", "--overwrite"])
    assert captured["kw"]["overwrite"] is True
    cli.main(["deck", "make", "--html", "x.html"])
    assert captured["kw"]["overwrite"] is False


def test_deck_make_overwrite_false_not_bool_true(monkeypatch, capsys):
    # P0-4 回归：--overwrite false 曾因 bool("false") is True 被翻成 True。
    from offipy import cli

    captured = {}

    def fake_make(html, **kw):
        captured["kw"] = kw
        return r"C:\out\deck.pptx"

    monkeypatch.setattr("offipy.deck.make", fake_make)
    cli.main(["deck", "make", "--html", "x.html", "--overwrite", "false"])
    assert captured["kw"]["overwrite"] is False
    cli.main(["deck", "make", "--html", "x.html", "--overwrite", "true"])
    assert captured["kw"]["overwrite"] is True


def test_deck_make_layouts_false(monkeypatch, capsys):
    # P0-4 回归：--layouts false 必须为 False，不能走 bool("false")。
    from offipy import cli

    captured = {}

    def fake_make(html, **kw):
        captured["kw"] = kw
        return r"C:\out\deck.pptx"

    monkeypatch.setattr("offipy.deck.make", fake_make)
    cli.main(["deck", "make", "--html", "x.html", "--layouts", "false"])
    assert captured["kw"]["apply_layouts"] is False
    cli.main(["deck", "make", "--html", "x.html", "--layouts"])
    assert captured["kw"]["apply_layouts"] is True


def test_deck_make_unknown_option_rejected(capsys):
    # P0-4：未知 --key 不再被 REMAINDER 吞掉，argparse 直接 exit 2。
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "make", "--html", "x.html", "--bogus", "1"])
    assert exc.value.code == 2
    assert "unrecognized arguments" in capsys.readouterr().err


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


def test_deck_outline_missing_input_raises():
    from offipy import cli

    with pytest.raises(SystemExit):
        cli.main(["deck", "outline"])


def test_deck_outline_missing_file_friendly_error(tmp_path, capsys):
    # 回归：--input 指向不存在的文件时曾裸 FileNotFoundError traceback，
    # 现在应 SystemExit 带友好提示。
    from offipy import cli

    missing = tmp_path / "nope.md"
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline", "--input", str(missing)])
    assert "找不到文件" in str(exc.value)


def test_deck_outline_bad_format_friendly_error(tmp_path, capsys):
    # 回归：非法大纲（缺 # 主标题）时曾裸 ValueError traceback，
    # 现在应 SystemExit 带友好提示。
    from offipy import cli

    bad = tmp_path / "bad.md"
    bad.write_text("## 只有页面\n- 无标题\n", encoding="utf-8")
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline", "--input", str(bad)])
    assert "大纲格式错误" in str(exc.value)


def test_check_subcommand_parses():
    args = build_parser().parse_args(["check"])
    assert args.app == "check"
    assert args.json is False


def test_check_subcommand_parses_json():
    args = build_parser().parse_args(["check", "--json"])
    assert args.app == "check"
    assert args.json is True


def test_check_dispatch_passes_json(monkeypatch):
    from offipy import cli

    captured = {}

    def fake_check_main(json_output=False):
        captured["json_output"] = json_output
        return 0

    monkeypatch.setattr("offipy.envcheck.main", fake_check_main)
    assert cli.main(["check"]) == 0
    assert captured["json_output"] is False
    assert cli.main(["check", "--json"]) == 0
    assert captured["json_output"] is True


def test_check_dispatch_propagates_exit_code(monkeypatch):
    from offipy import cli

    monkeypatch.setattr("offipy.envcheck.main", lambda json_output=False: 1)
    assert cli.main(["check"]) == 1


# --- 批次2：offipy server status/stop/restart ---


def test_server_subcommand_parses():
    args = build_parser().parse_args(["server"])
    assert args.app == "server"
    assert args.action == "status"  # 缺省 status
    assert build_parser().parse_args(["server", "stop"]).action == "stop"
    assert build_parser().parse_args(["server", "restart"]).action == "restart"


def test_server_status_dispatch(monkeypatch, capsys):
    from offipy import cli

    monkeypatch.setattr("offipy.cli.ensure_server", lambda: None)
    monkeypatch.setattr("offipy.cli.server_status", lambda: {"version": "0.9.0", "pid": 1})
    assert cli.main(["server", "status"]) is None
    assert "0.9.0" in capsys.readouterr().out


def test_server_status_not_running(monkeypatch, capsys):
    from offipy import cli

    # P0-3：未运行只报状态，不隐式拉起（不再调 ensure_server）
    called = []
    monkeypatch.setattr("offipy.cli.ensure_server", lambda: called.append("ensure"))
    monkeypatch.setattr("offipy.cli.server_status", lambda: None)
    assert cli.main(["server", "status"]) is None
    assert "未在运行" in capsys.readouterr().out
    assert called == []


def test_server_stop_dispatch(monkeypatch, capsys):
    from offipy import cli

    monkeypatch.setattr("offipy.cli.stop_server", lambda: "server 已停止")
    assert cli.main(["server", "stop"]) is None
    assert "已停止" in capsys.readouterr().out


def test_server_restart_dispatch(monkeypatch, capsys):
    from offipy import cli

    calls = []
    monkeypatch.setattr("offipy.cli.stop_server", lambda: calls.append("stop") or "server 已停止")
    monkeypatch.setattr("offipy.cli.ensure_server", lambda: calls.append("ensure"))
    assert cli.main(["server", "restart"]) is None
    assert calls == ["stop", "ensure"]
    assert "已重启" in capsys.readouterr().out


# --- 批次4：复杂参数系统 + 未知参数校验 ---


def test_parse_kwargs_repeat_flag_aggregates():
    from offipy.cli import _parse_kwargs

    # 值保留原始字符串（类型转换交给 _coerce_kwargs 按签名注解做）
    assert _parse_kwargs(["--lines", "a", "--lines", "b"]) == {"lines": ["a", "b"]}
    assert _parse_kwargs(["--n", "1", "--n", "2"]) == {"n": ["1", "2"]}


def test_parse_kwargs_payload_overrides():
    from offipy.cli import _parse_kwargs

    kw = _parse_kwargs(["--layout", "2", "--payload", '{"layout": 5, "nested": {"a": 1}}'])
    assert kw["layout"] == 5
    assert kw["nested"] == {"a": 1}


def test_parse_kwargs_payload_bad_json_exits():
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit):
        _parse_kwargs(["--payload", "{not json"])


def test_parse_kwargs_payload_must_be_object():
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit):
        _parse_kwargs(["--payload", "[1, 2]"])


def test_validate_kwargs_rejects_unknown_exit_2(monkeypatch, capsys):
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli._validate_kwargs("ppt", "add_slide", {"bogus": 1})
    assert exc.value.code == 2
    assert "unrecognized arguments" in capsys.readouterr().err


def test_validate_kwargs_accepts_known():
    from offipy import cli

    cli._validate_kwargs("ppt", "add_slide", {"layout": 2})
    cli._validate_kwargs("word", "add_list", {"lines": ["a", "b"], "style": "bullet"})


def test_validate_kwargs_unknown_op_deferred_to_server():
    from offipy import cli

    cli._validate_kwargs("ppt", "no_such_op", {"x": 1})  # 不抛：未知 op 交给 server


# --- 批次4：CLI 按签名类型转换 ---


def test_coerce_kwargs_str_keeps_string():
    from offipy.cli import _coerce_kwargs

    # "00123" 前导零保留、"true" 不当 bool
    assert _coerce_kwargs("word", "add_heading", {"text": "00123"}) == {"text": "00123"}
    assert _coerce_kwargs("ppt", "set_title", {"text": "true"}) == {"text": "true"}


def test_coerce_kwargs_int():
    from offipy.cli import _coerce_kwargs

    assert _coerce_kwargs("ppt", "add_slide", {"layout": "2"}) == {"layout": 2}


def test_coerce_kwargs_bool():
    from offipy.cli import _coerce_kwargs

    assert _coerce_kwargs("excel", "page_setup", {"center_horizontally": "true"}) == {
        "center_horizontally": True
    }
    assert _coerce_kwargs("excel", "page_setup", {"center_horizontally": "off"}) == {
        "center_horizontally": False
    }


def test_coerce_kwargs_int_invalid_exits(capsys):
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli._coerce_kwargs("ppt", "add_slide", {"layout": "abc"})
    assert exc.value.code == 2
    assert "整数" in capsys.readouterr().err


def test_coerce_kwargs_list_wraps_scalar():
    from offipy.cli import _coerce_kwargs

    assert _coerce_kwargs("word", "add_list", {"lines": "x"}) == {"lines": ["x"]}
    assert _coerce_kwargs("word", "add_list", {"lines": ["a", "b"]}) == {"lines": ["a", "b"]}


def test_coerce_kwargs_unknown_op_passthrough():
    from offipy.cli import _coerce_kwargs

    assert _coerce_kwargs("ppt", "no_such_op", {"x": "1"}) == {"x": "1"}


def test_main_coerces_before_call(monkeypatch):
    from offipy import cli

    captured = {}

    def fake_call(app, op, **kw):
        captured.update(kw)
        return None

    monkeypatch.setattr("offipy.cli.call", fake_call)
    cli.main(["ppt", "add_slide", "--layout", "2"])
    assert captured == {"layout": 2}
