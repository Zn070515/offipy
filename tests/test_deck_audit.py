"""deck 审计门禁：render_with_report 在 report/strict 模式下审计、替换/拒绝的正确性。

mock _render_tmp + audit_pptx（不跑真实转换/审计）：report 模式替换并返回
RenderResult；strict 达阈值抛 AuditGateError（report 可访问、tmp 无残留、
旧目标不动）；strict 通过才替换。
"""

import os
from contextlib import contextmanager
from pathlib import Path

import pytest

from offipy import deck
from offipy.audit import Severity
from offipy.deck import AuditGateError, RenderResult, render_with_report
from offipy.exceptions import InvalidArgumentError


def _fake_audit_factory(severity: Severity | None):
    """按 severity 返回 fake audit_pptx；None 表示报告但无 finding。"""

    def fake(path, config=None):
        class _Report:
            max_severity = severity

            def to_dict(self):
                return {"max_severity": severity.name if severity else None}

        return _Report()

    return fake


@pytest.fixture(autouse=True)
def _no_browser(monkeypatch):
    # 单测不真启动 chromium：render_with_report 的浏览器前置检查换成 no-op
    monkeypatch.setattr(deck, "_preflight_browser", lambda: None)


@pytest.fixture
def fake_tmp(monkeypatch, tmp_path):
    """mock _render_tmp：产出 (tmp_pptx, final_out)，finally 清理 tmp；记录 replace。"""
    real_replace = os.replace  # deck.os 即 os 模块；先存原实现避免递归
    calls = {"replaced": False, "final": None}

    @contextmanager
    def _fake_tmp(
        html, out, only_slides, no_visual_audit, timeout, theme, apply_layouts, overwrite, **kw
    ):
        final = str(tmp_path / "deck.pptx")
        tmp = str(tmp_path / ".deck.pptx")
        Path(tmp).write_bytes(b"tmp")
        try:
            yield tmp, final
        finally:
            if Path(tmp).exists():
                Path(tmp).unlink()

    monkeypatch.setattr(deck, "_render_tmp", _fake_tmp)

    def fake_replace(src, dst):
        calls["replaced"] = True
        calls["final"] = dst
        real_replace(src, dst)

    monkeypatch.setattr(deck.os, "replace", fake_replace)
    return calls


def test_report_mode_replaces_and_returns_render_result(tmp_path, monkeypatch, fake_tmp):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit_factory(None))

    result = render_with_report(str(html), audit_mode="report")

    assert isinstance(result, RenderResult)
    assert result.output_path.endswith("deck.pptx")
    assert result.to_dict()["output_path"] == result.output_path
    assert result.to_dict()["audit"]["max_severity"] is None
    assert fake_tmp["replaced"] is True
    assert fake_tmp["final"] == result.output_path
    assert Path(result.output_path).read_bytes() == b"tmp"  # 真实 replace 落盘


def test_strict_pass_replaces(tmp_path, monkeypatch, fake_tmp):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit_factory(Severity.LOW))

    result = render_with_report(str(html), audit_mode="strict", fail_on=Severity.HIGH)

    assert isinstance(result, RenderResult)
    assert fake_tmp["replaced"] is True
    assert result.audit_report.max_severity == Severity.LOW


def test_strict_hit_raises_gate_error_and_keeps_target(tmp_path, monkeypatch, fake_tmp):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit_factory(Severity.HIGH))

    with pytest.raises(AuditGateError) as exc:
        render_with_report(str(html), audit_mode="strict", fail_on=Severity.HIGH)

    assert exc.value.report.max_severity == Severity.HIGH
    assert exc.value.fail_on == Severity.HIGH
    assert fake_tmp["replaced"] is False  # 替换未发生
    assert not Path(tmp_path / ".deck.pptx").exists()  # tmp 已清理
    assert not Path(tmp_path / "deck.pptx").exists()  # 旧目标不动（无新文件落盘）


def test_strict_hit_preserves_existing_target(tmp_path, monkeypatch, fake_tmp):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    existing = tmp_path / "deck.pptx"
    existing.write_bytes(b"precious")
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit_factory(Severity.HIGH))

    with pytest.raises(AuditGateError):
        render_with_report(str(html), audit_mode="strict", fail_on=Severity.MID)

    assert existing.read_bytes() == b"precious"  # 已存在 .pptx 未被破坏


def test_report_mode_does_not_raise_even_when_gate_hit(tmp_path, monkeypatch, fake_tmp):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit_factory(Severity.HIGH))

    result = render_with_report(str(html), audit_mode="report", fail_on=Severity.LOW)

    assert isinstance(result, RenderResult)
    assert result.audit_report.max_severity == Severity.HIGH
    assert fake_tmp["replaced"] is True


# ---------------------------------------------------------------- 阻断项 2：audit_mode 校验


@pytest.mark.parametrize("bad_mode", ["strcit", "", "Strict", "report "])
def test_invalid_audit_mode_raises_before_render(tmp_path, fake_tmp, bad_mode):
    # 拼错 audit_mode 绝不静默退化绕 strict 门禁：进入渲染前抛 InvalidArgumentError
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    with pytest.raises(InvalidArgumentError):
        render_with_report(str(html), audit_mode=bad_mode)  # type: ignore[arg-type]
    assert fake_tmp["replaced"] is False  # 非法值绝不执行 os.replace


def test_non_severity_fail_on_raises_before_render(tmp_path, fake_tmp):
    html = tmp_path / "deck.html"
    html.write_text("<html><body>deck</body></html>", encoding="utf-8")
    with pytest.raises(InvalidArgumentError):
        render_with_report(str(html), audit_mode="strict", fail_on="HIGH")  # type: ignore[arg-type]
    assert fake_tmp["replaced"] is False
