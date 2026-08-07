"""CLI 参数解析测试（不依赖 Office）。"""

import pytest

from offipy.cli import build_parser
from offipy.exceptions import ComOperationError, FileConflictError, InvalidArgumentError


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


def test_quit_parser_has_force_flag():
    # P1-5：offipy quit <target> --force 强制退出既有 Office 实例
    args = build_parser().parse_args(["quit", "excel", "--force"])
    assert args.force is True
    args = build_parser().parse_args(["quit", "excel"])
    assert args.force is False


def test_office_subcommand_has_kwargs():
    args = build_parser().parse_args(["ppt", "add_slide", "--layout", "2"])
    assert args.app == "ppt"
    assert args.kwargs == ["--layout", "2"]


def test_parse_kwargs_flag_without_value():
    from offipy.cli import _parse_kwargs

    # 无值 flag → True；--key 归一化 - → _（no-open 等价 no_open，README 双写法一致）
    assert _parse_kwargs(["--no-open"]) == {"no_open": True}
    assert _parse_kwargs(["--no-open", "--out", "x.pptx"]) == {
        "no_open": True,
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
    # #29：deck 缺必填参数 → stderr + exit 2（对齐 argparse 参数错误语义）
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline"])
    assert exc.value.code == 2


def test_deck_make_missing_html_exits_2():
    # #29：deck make 缺 --html → stderr 用法提示 + exit 2（曾是 exit 1）
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "make"])
    assert exc.value.code == 2


def test_deck_make_audit_mode_missing_html_exits_2():
    # #29：--audit-mode 走 _deck_make_with_audit 分支，同样缺 --html → exit 2
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "make", "--audit-mode", "report"])
    assert exc.value.code == 2


def test_deck_outline_missing_file_friendly_error(tmp_path, capsys):
    # 回归：--input 指向不存在的文件时曾裸 FileNotFoundError traceback，
    # 现在应 SystemExit 带友好提示。S5：运行前非法输入 → exit 2（stderr 含路径）。
    from offipy import cli

    missing = tmp_path / "nope.md"
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline", "--input", str(missing)])
    assert exc.value.code == 2
    assert "找不到文件" in capsys.readouterr().err


def test_deck_outline_bad_format_friendly_error(tmp_path, capsys):
    # 回归：非法大纲（缺 # 主标题）时曾裸 ValueError traceback，
    # 现在应 SystemExit 带友好提示。S5：运行前非法输入 → exit 2。
    from offipy import cli

    bad = tmp_path / "bad.md"
    bad.write_text("## 只有页面\n- 无标题\n", encoding="utf-8")
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline", "--input", str(bad)])
    assert exc.value.code == 2
    assert "大纲格式错误" in capsys.readouterr().err


def test_deck_outline_out_unwritable_dir_exits_2(tmp_path, capsys):
    # S5：--out 指向不存在目录下的路径 → 曾裸 FileNotFoundError traceback，
    # 现在 catch OSError → 友好消息 + exit 2（输出路径非法）。
    from offipy import cli

    md = tmp_path / "outline.md"
    md.write_text("# T\n\n## 页\n- 甲\n", encoding="utf-8")
    bad_out = tmp_path / "nope_dir" / "deck.html"
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline", "--input", str(md), "--out", str(bad_out)])
    assert exc.value.code == 2
    assert "无法写入输出文件" in capsys.readouterr().err


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

    def fake_check_main(json_output=False, profile=None):
        captured["json_output"] = json_output
        captured["profile"] = profile
        return 0

    monkeypatch.setattr("offipy.envcheck.main", fake_check_main)
    assert cli.main(["check"]) == 0
    assert captured["json_output"] is False
    assert captured["profile"] is None
    assert cli.main(["check", "--json"]) == 0
    assert captured["json_output"] is True


def test_check_dispatch_passes_profile(monkeypatch):
    from offipy import cli

    captured = {}

    def fake_check_main(json_output=False, profile=None):
        captured["profile"] = profile
        return 0

    monkeypatch.setattr("offipy.envcheck.main", fake_check_main)
    assert cli.main(["check", "--profile", "office"]) == 0
    assert captured["profile"] == "office"
    assert cli.main(["check", "--profile", "deck"]) == 0
    assert captured["profile"] == "deck"


def test_check_dispatch_propagates_exit_code(monkeypatch):
    from offipy import cli

    monkeypatch.setattr("offipy.envcheck.main", lambda json_output=False, profile=None: 1)
    assert cli.main(["check"]) == 1


# --- 批次2：offipy server status/stop/restart ---


def test_server_subcommand_parses():
    args = build_parser().parse_args(["server"])
    assert args.app == "server"
    assert args.action == "status"  # 缺省 status
    assert build_parser().parse_args(["server", "stop"]).action == "stop"
    assert build_parser().parse_args(["server", "restart"]).action == "restart"


def test_server_subcommand_port_inherits_parent():
    # §11 回归：子解析器 --port 默认 None 曾覆盖顶层值，`--port 8891 server status`
    # 丢掉端口；SUPPRESS 后未给时不覆盖父级值。
    args = build_parser().parse_args(["--port", "8891", "server", "status"])
    assert args.port == 8891
    assert build_parser().parse_args(["server", "status"]).port is None  # 顶层缺省


def test_log_subcommand_port_inherits_or_overrides():
    args = build_parser().parse_args(["--port", "8891", "log"])
    assert args.port == 8891  # 顶层传入，子命令继承
    args = build_parser().parse_args(["log", "--port", "8892"])
    assert args.port == 8892  # 子命令显式传入覆盖


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
    # S5：#29 未清完——--payload JSON 解析失败原 SystemExit(str)→exit 1，
    # 现统一走 _usage_exit → exit 2（用法错误）。
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit) as exc:
        _parse_kwargs(["--payload", "{not json"])
    assert exc.value.code == 2


def test_parse_kwargs_payload_must_be_object():
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit):
        _parse_kwargs(["--payload", "[1, 2]"])


def test_parse_kwargs_dash_to_underscore():
    # 传输层参数之外：--doc-id 归一化为 doc_id（README 双写法不再打架）
    from offipy.cli import _parse_kwargs

    assert _parse_kwargs(["--doc-id", "book1"]) == {"doc_id": "book1"}
    assert _parse_kwargs(["--slide-idx", "2"]) == {"slide_idx": "2"}


def test_payload_array_error_hints_list_usage(capsys):
    # --payload 传数组时错误信息要提示 list 参数的正确写法（重复 --key / key 数组）。
    # S5：消息走 _usage_exit 到 stderr，SystemExit 带 int 2。
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit) as exc:
        _parse_kwargs(["--payload", "[1, 2]"])
    assert exc.value.code == 2
    msg = capsys.readouterr().err
    assert "list" in msg
    assert "重复 --key" in msg


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


def test_coerce_kwargs_line_spacing_keeps_string():
    from offipy.cli import _coerce_kwargs

    # S4 Task 3：CLI 的 --line_spacing 1.5 必须保留字符串（str|float 联合不做
    # 数值转换——line_spacing 语义上既接受字符串枚举也接受数值）
    assert _coerce_kwargs("word", "format_paragraph", {"line_spacing": "1.5"}) == {
        "line_spacing": "1.5"
    }
    assert _coerce_kwargs("word", "format_paragraph", {"line_spacing": "single"}) == {
        "line_spacing": "single"
    }


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
    # add_slide 是破坏性 op：必须带目标（doc_id/expected-target/follow-active）
    cli.main(["ppt", "add_slide", "--layout", "2", "--follow-active"])
    assert captured == {"layout": 2, "follow_active": True}


# --- 批次6：必填参数预校验（不碰 COM）+ mcp 缺 extra 友好报错 ---


def test_required_arg_missing_exits_2(capsys):
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli._validate_required("excel", "set_cell", {"cell": "A1"})
    assert exc.value.code == 2
    err = capsys.readouterr().err
    assert "缺少必填参数" in err
    assert "--sheet" in err and "--value" in err


def test_required_arg_present_passes():
    from offipy import cli

    cli._validate_required("excel", "set_cell", {"sheet": 1, "cell": "A1", "value": 5})


def test_required_unknown_op_deferred():
    # 未知 op 交给 server 侧报错，CLI 不预拦
    from offipy import cli

    cli._validate_required("ppt", "no_such_op", {})


def test_main_missing_required_does_not_call(monkeypatch, capsys):
    # 必填缺失 → 预校验 exit 2，server 不该被调用（不拉起、不碰 COM）
    from offipy import cli

    def boom(*a, **kw):
        raise AssertionError("必填缺失时不该调 server")

    monkeypatch.setattr("offipy.cli.call", boom)
    with pytest.raises(SystemExit) as exc:
        cli.main(["excel", "set_cell", "--cell", "A1"])
    assert exc.value.code == 2
    assert "缺少必填参数" in capsys.readouterr().err


def test_required_payload_keys_count_as_provided(monkeypatch):
    # --payload 注入的键也算已提供：必填齐全则放行到 call
    from offipy import cli

    captured = {}

    def fake_call(app, op, **kw):
        captured.update(kw)
        return None

    monkeypatch.setattr("offipy.cli.call", fake_call)
    # set_cell 是破坏性 op：payload 之外补 --follow-active 满足目标要求
    cli.main(
        [
            "excel",
            "set_cell",
            "--payload",
            '{"sheet": 1, "cell": "A1", "value": 5}',
            "--follow-active",
        ]
    )
    assert captured == {"sheet": 1, "cell": "A1", "value": 5, "follow_active": True}


# --- 批次 B1：传输层参数 expected-target / follow-active（P0-1/P0-3） ---


def test_parse_kwargs_expected_target_json():
    from offipy.cli import _parse_kwargs

    kw = _parse_kwargs(["--expected-target", '{"doc_id": "book1"}'])
    assert kw == {"expected_target": {"doc_id": "book1"}}


def test_parse_kwargs_expected_target_bad_json_exits():
    # S5：#29 未清完——--expected-target JSON 解析失败原 exit 1，现统一 exit 2。
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit) as exc:
        _parse_kwargs(["--expected-target", "{not json"])
    assert exc.value.code == 2


def test_parse_kwargs_expected_target_must_be_object():
    from offipy.cli import _parse_kwargs

    with pytest.raises(SystemExit):
        _parse_kwargs(["--expected-target", "[1, 2]"])


def test_parse_kwargs_follow_active_flag():
    from offipy.cli import _parse_kwargs

    assert _parse_kwargs(["--follow-active"]) == {"follow_active": True}
    assert _parse_kwargs(["--follow-active", "false"]) == {"follow_active": False}


def test_validate_kwargs_accepts_transport_params():
    from offipy.cli import _validate_kwargs

    # 破坏性 op 放行 expected_target / follow_active
    _validate_kwargs("excel", "set_cell", {"expected_target": {"doc_id": "b1"}})
    _validate_kwargs("excel", "set_cell", {"follow_active": True})


def test_validate_kwargs_rejects_transport_params_on_readonly():
    # 只读 op 不允许 expected_target（传输层绑定对只读 op 无意义）
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli._validate_kwargs("excel", "read_range", {"expected_target": {"doc_id": "b1"}})
    assert exc.value.code == 2


def test_validate_destructive_target_missing_exits_2(capsys):
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli._validate_destructive_target(
            "excel", "set_cell", {"sheet": 1, "cell": "A1", "value": 5}
        )
    assert exc.value.code == 2
    err = capsys.readouterr().err
    assert "必须显式指定目标文档" in err


def test_validate_destructive_target_doc_id_passes():
    from offipy.cli import _validate_destructive_target

    _validate_destructive_target("excel", "set_cell", {"doc_id": "b1"})
    _validate_destructive_target("excel", "set_cell", {"expected_target": {"doc_id": "b1"}})
    _validate_destructive_target("excel", "set_cell", {"follow_active": True})


def test_validate_destructive_target_skips_quit():
    from offipy.cli import _validate_destructive_target

    _validate_destructive_target("excel", "quit", {})  # quit 无 doc_id，不拦截


def test_validate_destructive_target_export_op_missing_exits_2(capsys):
    # P0-3：导出 op（requires_target）缺目标同样被拦截——防导出错文档
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli._validate_destructive_target("ppt", "export_slides", {"out_dir": "x"})
    assert exc.value.code == 2
    assert "必须显式指定目标文档" in capsys.readouterr().err


def test_validate_destructive_target_export_op_with_target_passes():
    from offipy.cli import _validate_destructive_target

    _validate_destructive_target("ppt", "export_slides", {"out_dir": "x", "doc_id": "p1"})
    _validate_destructive_target("ppt", "save_pdf", {"path": "x.pdf", "follow_active": True})
    _validate_destructive_target(
        "excel", "save_pdf", {"path": "x.pdf", "expected_target": {"doc_id": "b1"}}
    )


def test_main_destructive_without_target_exits_2(monkeypatch, capsys):
    # 破坏性 op 缺目标 → 调用前 exit 2，不拉起 server/碰 COM
    from offipy import cli

    def boom(*a, **kw):
        raise AssertionError("缺目标时不该调 server")

    monkeypatch.setattr("offipy.cli.call", boom)
    with pytest.raises(SystemExit) as exc:
        cli.main(["excel", "set_cell", "--sheet", "1", "--cell", "A1", "--value", "5"])
    assert exc.value.code == 2


def test_main_expected_target_passed_through(monkeypatch):
    from offipy import cli

    captured = {}

    def fake_call(app, op, **kw):
        captured.update(kw)
        return None

    monkeypatch.setattr("offipy.cli.call", fake_call)
    cli.main(
        [
            "excel",
            "set_cell",
            "--sheet",
            "1",
            "--cell",
            "A1",
            "--value",
            "5",
            "--expected-target",
            '{"doc_id": "book1"}',
        ]
    )
    assert captured["expected_target"] == {"doc_id": "book1"}
    assert "value" in captured  # value 按 schema 为 Any，原样透传


def test_mcp_missing_extra_friendly_error(monkeypatch, capsys):
    # 缺 mcp extra → 不再裸 ImportError traceback，exit 2 + 安装提示
    import sys

    from offipy import cli

    monkeypatch.setitem(sys.modules, "offipy.mcp_server", None)
    assert cli.main(["mcp"]) == 2
    err = capsys.readouterr().err
    assert "offipy[mcp]" in err


def test_mcp_present_runs(monkeypatch):
    from offipy import cli

    called = []

    def fake_main():
        called.append("mcp")

    monkeypatch.setattr("offipy.mcp_server.main", fake_main)
    assert cli.main(["mcp"]) is None
    assert called == ["mcp"]


# --- S5 Task 1：共享 CLI 边界错误处理 ---


def test_clean_error_message_strips_trace_frames_and_type_prefix():
    # S5：client._raise_error 拼出的消息带内部代码帧行与 "{TypeName}: " 前缀，
    # _clean_error_message 应去掉帧行（空白开头）与类型名前缀，保留 [app::op] 失败: 上下文。
    from offipy.cli import _clean_error_message

    msg = (
        "[word::open_doc] 失败: InvalidArgumentError: 源文件不存在: C:\\nope.docx\n"
        '    File "C:\\...\\word.py", line 292, in open_doc\n'
        "      raise InvalidArgumentError(...)\n"
        "  InvalidArgumentError: 源文件不存在: C:\\nope.docx"
    )
    assert _clean_error_message(msg) == "[word::open_doc] 失败: 源文件不存在: C:\\nope.docx"


def test_clean_error_message_keeps_non_type_message():
    # 无 TypeName 前缀的消息原样保留（不匹配 "{Type}: " 正则）
    from offipy.cli import _clean_error_message

    msg = "[ppt::export_slides] 失败: [WinError 183] 无法创建文件"
    assert _clean_error_message(msg) == msg


def test_clean_error_message_keeps_connection_failed_reason():
    # 连接失败消息不以 "[app::op] 失败: " 开头 → 类型名前缀不得被剥离
    from offipy.cli import _clean_error_message

    msg = "[excel::save] 连接失败: ConnectionRefusedError: [WinError 10061] 拒绝连接"
    assert _clean_error_message(msg) == msg


def test_clean_error_message_all_frame_lines_returns_empty():
    # 防御：所有行都是帧行（空白开头）时不能原样返回泄漏，应返回空串
    from offipy.cli import _clean_error_message

    msg = (
        '    File "C:\\...\\word.py", line 292, in open_doc\n      raise InvalidArgumentError(...)'
    )
    assert _clean_error_message(msg) == ""


def test_main_invalid_argument_exits_2(monkeypatch, capsys):
    # S5 核心规则：CLI 边界 InvalidArgumentError → exit 2，且 stderr 简洁无 Traceback
    from offipy import cli
    from offipy.exceptions import InvalidArgumentError

    def boom(app, op, **kw):
        raise InvalidArgumentError(
            "[word::open_doc] 失败: InvalidArgumentError: 源文件不存在: C:\\nope.docx"
        )

    monkeypatch.setattr("offipy.cli.call", boom)
    assert cli.main(["word", "open_doc", "--path", "C:\\nope.docx"]) == 2
    err = capsys.readouterr().err
    assert "源文件不存在" in err
    assert "Traceback" not in err


def test_main_other_offipy_error_exits_1(monkeypatch, capsys):
    # 其余 OffipyError（运行时领域失败）→ exit 1，消息同样清洗
    from offipy import cli
    from offipy.exceptions import TargetNotFoundError

    def boom(app, op, **kw):
        raise TargetNotFoundError("[excel::save] 失败: TargetNotFoundError: 没有活动工作簿")

    monkeypatch.setattr("offipy.cli.call", boom)
    assert cli.main(["excel", "save", "--path", "x.xlsx", "--follow-active"]) == 1
    err = capsys.readouterr().err
    assert "没有活动工作簿" in err
    assert "Traceback" not in err


def test_ensure_writable_missing_parent_dir_raises(tmp_path):
    # S5 差距 4：save/save_pdf 父目录不存在 → pre-check InvalidArgumentError（exit 2 侧）
    from offipy.exceptions import InvalidArgumentError
    from offipy.paths import ensure_writable

    with pytest.raises(InvalidArgumentError, match="输出目录不存在"):
        ensure_writable(r"C:\nope_dir\x.docx")
    # 父路径是文件（非目录）→ 同样 InvalidArgumentError（isdir 覆盖）
    f = tmp_path / "afile"
    f.write_text("x", encoding="utf-8")
    with pytest.raises(InvalidArgumentError, match="输出目录不存在"):
        ensure_writable(str(f / "x.docx"))
    # 裸文件名（无父目录 → 用 CWD）不抛
    assert ensure_writable("bare.docx")
    # tmp_path 存在 → 不抛
    assert ensure_writable(str(tmp_path / "x.docx"))


# --- S5 Task 2：Word CLI 错误契约对齐 ---


def _assert_cli_error(monkeypatch, capsys, args, exc_cls, msg, code, *snippets, absent=()):
    """断言 cli.main 抛出的领域异常按 CLI 契约落 exit code + 清洗 stderr。

    mock offipy.cli.call 抛 exc_cls(msg)，验证边界把异常翻译成 exit code 2/1、
    stderr 含 snippets（op/路径/文案）、无 Traceback、且 absent 缺席
    （如内部代码帧行、类型名前缀）。Task 2 的三类异常都接受纯 message 字符串。
    """
    from offipy import cli

    def boom(app, op, **kw):
        raise exc_cls(msg)

    monkeypatch.setattr("offipy.cli.call", boom)
    assert cli.main(args) == code
    err = capsys.readouterr().err
    for s in snippets:
        assert s in err
    assert "Traceback" not in err
    for a in absent:
        assert a not in err


def test_word_open_doc_missing_file_exits_2(monkeypatch, capsys):
    # S5：word open_doc 运行前非法输入（源文件不存在）→ exit 2，stderr 含路径无 Traceback
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "open_doc", "--path", "C:\\nope.docx"],
        InvalidArgumentError,
        "[word::open_doc] 失败: InvalidArgumentError: 源文件不存在: C:\\nope.docx",
        2,
        "源文件不存在",
        "C:\\nope.docx",
        "open_doc",
    )


def test_word_open_doc_com_failure_exits_1(monkeypatch, capsys):
    # S5：word open_doc 运行时 COM 失败 → exit 1；stderr 清洗掉内部帧行与类型名前缀
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "open_doc", "--path", "C:\\notes.txt"],
        ComOperationError,
        "[word::open_doc] 失败: ComOperationError: 无法打开文件（非法格式）: C:\\notes.txt\n"
        '    File "C:\\...\\word.py", line 293, in open_doc\n'
        "      return self._register(self.app.Documents.Open(path))\n"
        "  ComOperationError: 无法打开文件（非法格式）: C:\\notes.txt",
        1,
        "无法打开文件",
        "C:\\notes.txt",
        "open_doc",
        absent=("word.py", "ComOperationError"),
    )


def test_word_save_conflict_exits_1(monkeypatch, capsys):
    # S5：word save 目标已存在未 overwrite → FileConflictError → exit 1
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "save", "--path", "C:\\out.docx", "--follow-active"],
        FileConflictError,
        "[word::save] 失败: FileConflictError: 目标文件已存在: C:\\out.docx"
        "（如确要覆盖请传 overwrite=True）",
        1,
        "目标文件已存在",
        "C:\\out.docx",
        "save",
    )


def test_word_save_missing_parent_dir_exits_2(monkeypatch, capsys):
    # S5：word save 父目录不存在 → InvalidArgumentError（运行前非法输出路径）→ exit 2
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "save", "--path", "C:\\nope_dir\\out.docx", "--follow-active"],
        InvalidArgumentError,
        "[word::save] 失败: InvalidArgumentError: 输出目录不存在: C:\\nope_dir",
        2,
        "输出目录不存在",
        "C:\\nope_dir",
        "save",
    )


def test_word_save_pdf_conflict_exits_1(monkeypatch, capsys):
    # S5：word save_pdf 目标已存在未 overwrite → FileConflictError → exit 1
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "save_pdf", "--path", "C:\\out.pdf", "--follow-active"],
        FileConflictError,
        "[word::save_pdf] 失败: FileConflictError: 目标文件已存在: C:\\out.pdf"
        "（如确要覆盖请传 overwrite=True）",
        1,
        "目标文件已存在",
        "C:\\out.pdf",
        "save_pdf",
    )


def test_word_save_pdf_missing_parent_dir_exits_2(monkeypatch, capsys):
    # S5：word save_pdf 父目录不存在 → InvalidArgumentError → exit 2
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "save_pdf", "--path", "C:\\nope_dir\\out.pdf", "--follow-active"],
        InvalidArgumentError,
        "[word::save_pdf] 失败: InvalidArgumentError: 输出目录不存在: C:\\nope_dir",
        2,
        "输出目录不存在",
        "C:\\nope_dir",
        "save_pdf",
    )


def test_word_insert_image_missing_file_exits_2(monkeypatch, capsys):
    # S5：word insert_image 源图片不存在 → InvalidArgumentError → exit 2
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["word", "insert_image", "--path", "C:\\nope.png", "--follow-active"],
        InvalidArgumentError,
        "[word::insert_image] 失败: InvalidArgumentError: 源文件不存在: C:\\nope.png",
        2,
        "源文件不存在",
        "C:\\nope.png",
        "insert_image",
    )


# --- S5 Task 3：Excel CLI 错误契约对齐 ---


def test_excel_open_book_missing_file_exits_2(monkeypatch, capsys):
    # S5：excel open_book 运行前非法输入（源文件不存在）→ exit 2，stderr 含 op/路径无 Traceback
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["excel", "open_book", "--path", "C:\\nope.xlsx"],
        InvalidArgumentError,
        "[excel::open_book] 失败: InvalidArgumentError: 源文件不存在: C:\\nope.xlsx",
        2,
        "excel",
        "open_book",
        "源文件不存在",
        "C:\\nope.xlsx",
    )


def test_excel_open_book_com_failure_exits_1(monkeypatch, capsys):
    # S5：excel open_book 运行时 COM 失败（非法格式）→ exit 1；
    # stderr 清洗掉内部代码帧行（excel.py）与类型名前缀（ComOperationError）
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["excel", "open_book", "--path", "C:\\notes.txt"],
        ComOperationError,
        "[excel::open_book] 失败: ComOperationError: 无法打开文件（非法格式）: C:\\notes.txt\n"
        '    File "C:\\...\\excel.py", line 254, in open_book\n'
        "      return self._register(self.app.Workbooks.Open(path))\n"
        "  ComOperationError: 无法打开文件（非法格式）: C:\\notes.txt",
        1,
        "excel",
        "open_book",
        "无法打开",
        "C:\\notes.txt",
        absent=("excel.py", "ComOperationError"),
    )


def test_excel_save_conflict_exits_1(monkeypatch, capsys):
    # S5：excel save 目标已存在未 overwrite → FileConflictError（运行时领域失败）→ exit 1
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["excel", "save", "--path", "C:\\out.xlsx", "--follow-active"],
        FileConflictError,
        "[excel::save] 失败: FileConflictError: 目标文件已存在: C:\\out.xlsx"
        "（如确要覆盖请传 overwrite=True）",
        1,
        "save",
        "目标文件已存在",
        "C:\\out.xlsx",
    )


def test_excel_save_missing_parent_dir_exits_2(monkeypatch, capsys):
    # S5：excel save 父目录不存在 → InvalidArgumentError（运行前非法输出路径）→ exit 2
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["excel", "save", "--path", "C:\\nope_dir\\out.xlsx", "--follow-active"],
        InvalidArgumentError,
        "[excel::save] 失败: InvalidArgumentError: 输出目录不存在: C:\\nope_dir",
        2,
        "save",
        "输出目录不存在",
        "C:\\nope_dir",
    )


def test_excel_save_pdf_conflict_exits_1(monkeypatch, capsys):
    # S5：excel save_pdf 目标已存在未 overwrite → FileConflictError → exit 1
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["excel", "save_pdf", "--path", "C:\\out.pdf", "--follow-active"],
        FileConflictError,
        "[excel::save_pdf] 失败: FileConflictError: 目标文件已存在: C:\\out.pdf"
        "（如确要覆盖请传 overwrite=True）",
        1,
        "save_pdf",
        "目标文件已存在",
        "C:\\out.pdf",
    )


def test_excel_save_pdf_missing_parent_dir_exits_2(monkeypatch, capsys):
    # S5：excel save_pdf 父目录不存在 → InvalidArgumentError → exit 2
    _assert_cli_error(
        monkeypatch,
        capsys,
        ["excel", "save_pdf", "--path", "C:\\nope_dir\\out.pdf", "--follow-active"],
        InvalidArgumentError,
        "[excel::save_pdf] 失败: InvalidArgumentError: 输出目录不存在: C:\\nope_dir",
        2,
        "save_pdf",
        "输出目录不存在",
        "C:\\nope_dir",
    )
