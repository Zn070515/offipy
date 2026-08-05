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


def test_render_atomic_cleans_orphan_audit_dir(tmp_path, monkeypatch):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    created = {}
    monkeypatch.setattr(deck.subprocess, "run", _fake_run_with_audit(created))

    pptx = deck.render(str(html), overwrite=True)
    # convert 的 <tmp>_audit 审计目录在 tmp 被替换后成为孤儿，render 应清理
    orphans = [p for p in tmp_path.iterdir() if p.name.endswith("_audit")]
    assert orphans == []
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
    with pytest.raises(RuntimeError):
        deck.render(str(html), overwrite=True)
    assert existing.read_bytes() == b"precious"
    assert _hidden_pptx(tmp_path) == []


# --- §7：mkstemp 临时文件（并发安全）+ 源缺失映射领域异常 ---


def test_render_missing_source_raises_invalid_argument(tmp_path):
    # 源 HTML 缺失 → offipy 领域异常，不再裸抛 FileNotFoundError
    with pytest.raises(deck.InvalidArgumentError) as exc:
        deck.render(str(tmp_path / "nope.html"))
    assert "nope.html" in str(exc.value)


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
