"""deck 管线测试：theme 注入（不跑真实转换，拦截子进程）。"""

import os
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

from offipy import deck


@pytest.fixture(autouse=True)
def _no_browser(monkeypatch):
    # 单测不真启动 chromium：render 的浏览器前置检查换成 no-op
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)


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


def test_render_apply_layouts_injects_layout_css(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><head></head><body><section class="slide" data-pptx-slide '
        'data-layout="cards-3">x</section></body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.render(str(html), apply_layouts=True)
    injected = created["cmd"][2]
    assert injected.endswith(".audited.html")
    assert not os.path.exists(injected), "注入副本应已清理"
    assert ".cards-3 .cards" in created["injected_content"]
    assert pptx.endswith("deck.pptx")


def test_render_theme_and_layouts_together(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><head><style data-theme="mckinsey"></style></head><body>'
        '<section class="slide" data-pptx-slide data-layout="cards-3">x</section></body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    deck.render(str(html), theme="mckinsey", apply_layouts=True)
    content = created["injected_content"]
    assert "--accent: #2251FF;" in content  # 主题 token
    assert ".cards-3 .cards" in content  # 布局 CSS
    assert "data-layouts" in content


def test_render_apply_layouts_false_unchanged(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><body><section class="slide" data-pptx-slide '
        'data-layout="cards-3">x</section></body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    deck.render(str(html))  # 默认不注入
    assert created["cmd"][2] == str(html)
    assert ".cards-3 .cards" not in created["injected_content"]  # 布局 CSS 未注入


# --- overwrite 覆盖保护（P1-5）：不静默覆盖已有 .pptx ---


def test_render_existing_out_refuses_without_overwrite(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    (tmp_path / "deck.pptx").write_bytes(b"existing")
    created = {}
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    with pytest.raises(FileExistsError):
        deck.render(str(html))
    # fail-fast：未跑转换
    assert "cmd" not in created


def test_render_overwrite_true_allows(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    (tmp_path / "deck.pptx").write_bytes(b"existing")
    created = {}
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.render(str(html), overwrite=True)
    assert pptx.endswith("deck.pptx")


def test_make_existing_out_refuses_without_overwrite(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    (tmp_path / "deck.pptx").write_bytes(b"existing")
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)

    with pytest.raises(FileExistsError):
        deck.make(str(html), open_live_flag=False)


def test_make_passes_overwrite_to_render(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    (tmp_path / "deck.pptx").write_bytes(b"existing")
    created = {}
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.make(str(html), open_live_flag=False, overwrite=True)
    assert pptx.endswith("deck.pptx")


def test_make_feedback_binds_export_to_rendered_doc(tmp_path, monkeypatch):
    # P0-2：feedback_dir 导出必须绑定本次渲染的 deck（open_live 返回的 doc_id），
    # overwrite 透传；绝不依赖「当前活动焦点」（防中途切到别的文稿）。
    calls = {}

    def fake_render(html, out=None, **kw):
        return str(tmp_path / "deck.pptx")

    def fake_open_live(pptx):
        calls["open_pptx"] = pptx
        return "pres7"

    def fake_export_slides(out_dir, width=1920, height=1080, doc_id=None, overwrite=False):
        calls["export"] = {"out_dir": out_dir, "doc_id": doc_id, "overwrite": overwrite}
        return []

    monkeypatch.setattr(deck, "render", fake_render)
    monkeypatch.setattr(deck, "open_live", fake_open_live)
    monkeypatch.setattr(deck, "export_slides", fake_export_slides)

    feedback = tmp_path / "fb"
    deck.make(
        str(tmp_path / "deck.html"),
        open_live_flag=False,
        feedback_dir=str(feedback),
        overwrite=True,
    )
    assert calls["open_pptx"] == str(tmp_path / "deck.pptx")
    assert calls["export"]["doc_id"] == "pres7"
    assert calls["export"]["overwrite"] is True
    assert calls["export"]["out_dir"] == str(feedback)


# --- P0-6 原子替换：失败不破坏已存在 .pptx，临时文件清理 ---


def _hidden_pptx(tmp_path):
    """render 留下的隐藏临时 .pptx（成功替换后应一个不剩）。"""
    return [
        p.name for p in tmp_path.iterdir() if p.name.startswith(".") and p.name.endswith(".pptx")
    ]


def _fake_run_with_audit(created):
    """mock subprocess.run：按 --out 产出 tmp .pptx 之外，还建 <tmp>_audit 审计目录。"""

    def fake_run(cmd, **kw):
        created["cmd"] = cmd
        out = cmd[cmd.index("--out") + 1]
        Path(out).write_bytes(b"fake pptx")
        audit = Path(out).with_name(Path(out).stem + "_audit")
        (audit / "_cache").mkdir(parents=True, exist_ok=True)
        (audit / "_cache" / "measurements.json").write_text("{}", encoding="utf-8")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    return fake_run


def test_render_atomic_success_replaces_and_cleans(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    (tmp_path / "deck.pptx").write_bytes(b"old content")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.render(str(html), overwrite=True)
    assert Path(pptx).read_bytes() == b"fake pptx"  # 新内容到位
    assert _hidden_pptx(tmp_path) == []  # 临时 .pptx 已清理


def test_render_atomic_preserves_audit_dir_under_final_name(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_with_audit(created))

    pptx = deck.render(str(html), overwrite=True)
    # #11：convert 的 <tmp>_audit 改名为 <final>_audit 保留（aesthetic/feedback
    # 按 final stem 自动发现 measurements），不留 <tmp>_audit 孤儿
    assert (tmp_path / "deck_audit" / "_cache" / "measurements.json").exists()
    audits = [p.name for p in tmp_path.iterdir() if p.name.endswith("_audit")]
    assert audits == ["deck_audit"]
    assert Path(pptx).read_bytes() == b"fake pptx"


def test_render_atomic_failure_cleans_orphan_audit_dir(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    existing = tmp_path / "deck.pptx"
    existing.write_bytes(b"precious")

    def fail_run_with_audit(cmd, **kw):
        out = cmd[cmd.index("--out") + 1]
        audit = Path(out).with_name(Path(out).stem + "_audit")
        audit.mkdir(parents=True, exist_ok=True)
        return SimpleNamespace(returncode=1, stdout="oops", stderr="boom")

    monkeypatch.setattr(deck.subprocess, "run", fail_run_with_audit)
    with pytest.raises(deck.ConversionError):
        deck.render(str(html), overwrite=True)
    assert existing.read_bytes() == b"precious"  # 已存在 .pptx 未被破坏
    orphans = [p for p in tmp_path.iterdir() if p.name.endswith("_audit")]
    assert orphans == []


def test_render_atomic_convert_failure_preserves_existing(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    existing = tmp_path / "deck.pptx"
    existing.write_bytes(b"precious")
    created = {}

    def fail_run(cmd, **kw):
        created["cmd"] = cmd
        return SimpleNamespace(returncode=1, stdout="oops", stderr="boom")

    monkeypatch.setattr(deck.subprocess, "run", fail_run)
    with pytest.raises(deck.ConversionError):
        deck.render(str(html), overwrite=True)
    assert existing.read_bytes() == b"precious"  # 已存在 .pptx 未被破坏
    assert _hidden_pptx(tmp_path) == []


def test_render_atomic_postprocess_failure_preserves_existing(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    existing = tmp_path / "deck.pptx"
    existing.write_bytes(b"precious")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    def boom(html_path, pptx_path):
        raise RuntimeError("图表后处理失败")

    monkeypatch.setattr("offipy.charts.postprocess_charts", boom)
    # 后处理异常统一包装：RuntimeError → ConversionError（保留 __cause__）
    with pytest.raises(deck.ConversionError):
        deck.render(str(html), overwrite=True)
    assert existing.read_bytes() == b"precious"
    assert _hidden_pptx(tmp_path) == []


# --- §7：mkstemp 临时文件（并发安全）+ 源缺失映射领域异常 ---


def test_render_missing_source_raises_invalid_argument(tmp_path):
    # 源 HTML 缺失 → offipy 领域异常，不再裸抛 FileNotFoundError
    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(tmp_path / "nope.html"))
    assert "nope.html" in str(exc.value)


# --- B7：注入副本挪 TemporaryDirectory（不再污染源目录）+ 后处理异常包装 + 前置校验 ---


def test_render_theme_tmp_not_in_source_dir(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><head><style data-theme="mckinsey"></style></head>'
        '<body><section class="slide" data-pptx-slide>hi</section></body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    deck.render(str(html), theme="mckinsey")
    injected = created["cmd"][2]
    assert injected.endswith(".audited.html")
    assert not os.path.exists(injected), "注入副本应已清理"
    # 注入副本不再落在源目录（TemporaryDirectory），源目录只剩原 html + 产物
    assert set(p.name for p in tmp_path.iterdir()) == {"deck.html", "deck.pptx"}


def test_render_postprocess_valueerror_maps_to_invalid_argument(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    def boom(html_path, pptx_path):
        raise ValueError("图表数据 categories 必须是非空字符串列表")

    monkeypatch.setattr("offipy.charts.postprocess_charts", boom)
    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(html), overwrite=True)
    assert "图表" in str(exc.value)
    assert isinstance(exc.value.__cause__, ValueError)  # __cause__ 保留


def test_render_postprocess_runtimeerror_maps_to_conversion(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    def boom(html_path, pptx_path):
        raise RuntimeError("找不到 convert 审计产物")

    monkeypatch.setattr("offipy.charts.postprocess_charts", boom)
    with pytest.raises(deck.ConversionError) as exc:
        deck.render(str(html), overwrite=True)
    assert "图表" in str(exc.value)
    assert isinstance(exc.value.__cause__, RuntimeError)


def test_render_postprocess_icons_valueerror_maps(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    def boom(html_path, pptx_path):
        raise ValueError("图标数据非法")

    monkeypatch.setattr("offipy.icons.postprocess_icons", boom)
    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(html), overwrite=True)
    assert "图标" in str(exc.value)
    assert isinstance(exc.value.__cause__, ValueError)


def test_render_no_visual_audit_with_chart_rejected(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><body><section class="slide" data-pptx-slide><div class="chart" '
        'data-chart="bar" data-chart-data=\'{"categories":["a"],"series":'
        '[{"name":"s","values":[1]}]}\'></div></section></body></html>',
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(html), no_visual_audit=True)
    assert "data-chart" in str(exc.value)
    assert "cmd" not in created  # fail-fast：未触发转换


def test_render_no_visual_audit_with_icon_rejected(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text(
        '<html><body><section class="slide" data-pptx-slide>'
        '<svg class="icon" data-icon="ph:check" viewBox="0 0 256 256"></svg>'
        "</section></body></html>",
        encoding="utf-8",
    )
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(html), no_visual_audit=True)
    assert "data-icon" in str(exc.value)
    assert "cmd" not in created


def test_render_no_visual_audit_without_charts_ok(tmp_path, monkeypatch):
    # 回归：无图表/图标的纯 deck 配 no_visual_audit 不受前置校验影响
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run(created))

    pptx = deck.render(str(html), no_visual_audit=True)
    assert pptx.endswith("deck.pptx")
    assert created["cmd"][-1] == "--no-visual-audit"  # 转换器确实收到该开关


def test_render_concurrent_same_output_no_clash(tmp_path, monkeypatch):
    # 旧确定性临时名（.deck.tmp.pptx）并发渲染同一输出时会互踩：两个线程写
    # 同一个临时文件，一方的 finally 会删掉另一方还在用的文件。mkstemp 随机名
    # 保证两次转换各用各的临时文件。这里故意让第二个并发者转换失败（returncode
    # 1），只验证临时文件唯一性与清理——不制造两个线程同刻 os.replace 同一目标
    # 的竞争（Windows 上双写同一目标瞬时锁冲突是固有行为，非本回归点）。
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    seen = []
    lock = threading.Lock()

    def mock_run(cmd, **kw):
        out = cmd[cmd.index("--out") + 1]
        with lock:
            seen.append(out)
            idx = len(seen)
        time.sleep(0.1)  # 拉大并发窗口：两线程同时持有各自的 tmp
        if idx == 2:
            return SimpleNamespace(returncode=1, stdout="boom", stderr="")
        Path(out).write_bytes(b"fake pptx")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(deck.subprocess, "run", mock_run)
    results = {"err": None, "pptx": None}

    def do_render():
        try:
            results["pptx"] = deck.render(str(html), overwrite=True)
        except deck.ConversionError as e:
            results["err"] = str(e)

    t1 = threading.Thread(target=do_render)
    t2 = threading.Thread(target=do_render)
    t1.start()
    t2.start()
    t1.join(timeout=10)
    t2.join(timeout=10)
    assert len(seen) == 2
    assert len(set(seen)) == 2  # mkstemp 随机临时名，互不覆盖（旧确定性名会撞）
    assert results["err"] is not None  # 第二个并发者如实失败
    assert results["pptx"] is not None  # 第一个并发者正常完成
    assert Path(results["pptx"]).read_bytes() == b"fake pptx"
    assert _hidden_pptx(tmp_path) == []  # 两轮临时文件都清理干净


def test_render_concurrent_same_final_path_one_conflicts(tmp_path, monkeypatch):
    # review L605：并发渲染同一最终输出 → 先到者成功落盘，后到者的 fail-fast
    # preflight（overwrite=False 时 os.path.exists(final_out)）看到已存在输出 →
    # FileConflictError，绝不同时双写同一目标。
    # 线程命名 + Event 门控让交错确定性：B 的 preflight 阻塞到 A 完成 os.replace
    # 之后才判定，必然复现「一成功一冲突」；若 A 先于 B 到达 preflight（门未设），
    # B 直接看到已存在输出同样走冲突分支——两种时序都收敛到同一断言。
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    out = tmp_path / "deck.pptx"

    a_done = threading.Event()
    real_exists = os.path.exists

    def exists_gated(path):
        if os.path.abspath(path) == os.path.abspath(str(out)):
            if threading.current_thread().name == "render-B" and not a_done.is_set():
                a_done.wait(timeout=30)
            return real_exists(path)
        return real_exists(path)

    monkeypatch.setattr(os.path, "exists", exists_gated)

    def fake_run(cmd, **kw):
        out_arg = cmd[cmd.index("--out") + 1]
        Path(out_arg).write_bytes(b"fake pptx")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(deck.subprocess, "run", fake_run)

    results = {}

    def do_render_b():
        try:
            deck.render(str(html), out=str(out))
        except deck.FileConflictError as e:
            results["conflict"] = str(e)

    t2 = threading.Thread(target=do_render_b, name="render-B")
    t2.start()
    pptx = deck.render(str(html), out=str(out))  # A 主线程先完成落盘
    a_done.set()
    t2.join(timeout=30)

    assert pptx == os.path.abspath(str(out))
    assert Path(out).read_bytes() == b"fake pptx"
    assert "conflict" in results  # B 的 preflight 拒绝已存在输出
    assert _hidden_pptx(tmp_path) == []  # 双方临时文件都清理干净


# --- #22 open_live 文件锁：临时副本演示 + close_live 释放 + 可操作替换错误 ---


def _fake_call_returns(monkeypatch, doc_id):
    """mock call：记录最后调用，固定返回 doc_id。"""
    calls = {}

    def fake_call(app, op, **kw):
        calls["call"] = (app, op, kw)
        return doc_id

    monkeypatch.setattr(deck, "ensure_server", lambda: None)
    monkeypatch.setattr(deck, "call", fake_call)
    return calls


def test_open_live_presents_from_temp_copy_not_locking_source(tmp_path, monkeypatch):
    # #22：open_live 复制到 offipy-live-* 临时副本再让 PowerPoint 打开——锁在副本，
    # 源产物路径不被锁，同路径 re-render 不再 PermissionError。
    import tempfile

    src = tmp_path / "deck.pptx"
    src.write_bytes(b"fake pptx v1")
    calls = _fake_call_returns(monkeypatch, "pres9")

    doc_id = deck.open_live(str(src))
    assert doc_id == "pres9"
    app, op, kw = calls["call"]
    assert (app, op) == ("ppt", "open_pres")
    live = kw["path"]
    assert Path(live).is_file()
    assert Path(live).read_bytes() == b"fake pptx v1"
    assert Path(live).name.startswith("offipy-live-")
    assert Path(live).parent == Path(tempfile.gettempdir())
    assert src.exists(), "open_live 不应移动/删除源产物"
    assert deck._LIVE_TMP_PATHS.get("pres9") == live

    deck.close_live(doc_id)
    assert not Path(live).exists(), "open_live 测试未清理临时副本"
    assert doc_id not in deck._LIVE_TMP_PATHS


def test_close_live_closes_and_removes_temp_copy(tmp_path, monkeypatch):
    src = tmp_path / "deck.pptx"
    src.write_bytes(b"x")
    calls = _fake_call_returns(monkeypatch, "pres10")
    doc_id = deck.open_live(str(src))
    live = deck._LIVE_TMP_PATHS[doc_id]
    calls.clear()

    deck.close_live(doc_id)
    assert calls["call"] == ("ppt", "close_pres", {"doc_id": doc_id, "save": False})
    assert not Path(live).exists(), "close_live 未清理临时副本"
    assert doc_id not in deck._LIVE_TMP_PATHS


def test_render_same_path_works_after_open_live(tmp_path, monkeypatch):
    # #22 核心场景：open_live 后对同一输出路径再 render(overwrite=True) 不再被锁。
    src = tmp_path / "deck.pptx"
    src.write_bytes(b"v1")
    _fake_call_returns(monkeypatch, "pres11")
    deck.open_live(str(src))

    new = tmp_path / "new.pptx"
    new.write_bytes(b"v2")
    deck._atomic_replace(str(new), str(src))  # 模拟下次 render 的原子替换
    assert src.read_bytes() == b"v2"
    deck.close_live("pres11")


def test_atomic_replace_win32_locked_raises_actionable(tmp_path, monkeypatch):
    # #22：目标被 PowerPoint 锁定 → os.replace WinError 5 → ConversionError 带指引
    # （而非裸 PermissionError），CLI/用户能直接照做。
    from offipy.exceptions import ConversionError

    def locked_replace(src_, dst_):
        raise PermissionError(13, "Permission denied", dst_)

    monkeypatch.setattr(deck.os, "replace", locked_replace)
    with pytest.raises(ConversionError) as exc:
        deck._atomic_replace("tmp.pptx", str(tmp_path / "out.pptx"))
    assert "close_live" in str(exc.value)
    assert "out.pptx" in str(exc.value)
