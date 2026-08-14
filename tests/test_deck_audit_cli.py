"""offipy deck audit CLI 测试（Task 5 / S3）：源校验、单次分析、确定性输出、临时清理。

全部 mock，不碰 chromium / Office / 真实渲染。
"""

import json
import tempfile
import types
from pathlib import Path

import pytest

from offipy.art.models import (
    ArtFinding,
    ArtReport,
    ArtSlideReport,
    DeckQualityReport,
    DimensionAssessment,
)
from offipy.audit import Severity
from offipy.exceptions import ConversionError


def _finding(
    rule_id="art.hierarchy.title_size",
    dimension="hierarchy",
    severity=Severity.MID,
    message="标题字号过小（font_size_norm=0.018）。",
    slide_index=1,
    primary=None,
    details=None,
):
    return ArtFinding(
        rule_id=rule_id,
        dimension=dimension,
        severity=severity,
        message=message,
        confidence=0.8,
        slide_index=slide_index,
        primary=primary,
        details=details or {},
    )


def _report(*, with_art=True):
    """共享报告构建：一个页内 finding + 一个 deck 级 finding。"""
    if not with_art:
        return DeckQualityReport(geometry=None, art=None, warnings=[])
    art = ArtReport(
        profile="balanced",
        slides=[
            ArtSlideReport(
                slide_index=1,
                dimensions=[
                    DimensionAssessment(
                        dimension="hierarchy",
                        status="assessed",
                        grade="attention",
                        confidence=0.8,
                        findings=[_finding()],
                    )
                ],
            )
        ],
        deck_findings=[
            _finding(
                rule_id="art.consistency.suite",
                dimension="consistency",
                message="deck 级配色不一致",
                slide_index=None,
            )
        ],
    )
    return DeckQualityReport(geometry=None, art=art, warnings=[])


def _render_result():
    return types.SimpleNamespace(deck_quality=_report())


# ---------------------------------------------------------------- 源校验


def test_deck_audit_missing_source_exits_2():
    from offipy import cli

    # 无源 → usage 错误 exit 2
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "audit"])
    assert exc.value.code == 2

    # 同时给位置 source 与 --pptx → 互斥 exit 2
    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "audit", "--pptx", "a.pptx", "foo.html"])
    assert exc.value.code == 2


def test_deck_audit_make_rejects_positional_source():
    # 位置参数只属于 audit：make/outline 收到它 → 拒绝，不静默改变原命令语义
    from offipy import cli

    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "make", "foo.html"])
    assert exc.value.code == 2

    with pytest.raises(SystemExit) as exc:
        cli.main(["deck", "outline", "foo.md"])
    assert exc.value.code == 2


# ---------------------------------------------------------------- HTML 流


def test_deck_audit_html_calls_render_with_quality_report_once(monkeypatch, capsys):
    from offipy import cli

    calls = []

    def fake_render(html, **kw):
        calls.append((html, kw))
        return _render_result()

    monkeypatch.setattr("offipy.deck.render_with_quality_report", fake_render)
    cli.main(["deck", "audit", "x.html", "--theme", "mckinsey", "--layouts"])
    assert len(calls) == 1
    html, kw = calls[0]
    assert html == "x.html"
    assert kw["theme"] == "mckinsey"
    assert kw["apply_layouts"] is True
    assert kw["profile"] == "balanced"
    assert kw["pixel_analysis"] == "off"
    out = capsys.readouterr().out
    assert "offipy deck audit（profile=balanced）" in out
    assert "维度: hierarchy" in out
    assert "页 1 [MID]" in out
    assert "无自动建议" in out


def test_deck_audit_html_no_second_audit(monkeypatch, capsys):
    # HTML 流只允许一次分析：render_with_quality_report 内部已含几何+艺术，
    # CLI 绝不二次调 audit_pptx / build_scene / analyze_deck
    from offipy import cli

    calls = []

    def fake_render(html, **kw):
        calls.append(html)
        return _render_result()

    monkeypatch.setattr("offipy.deck.render_with_quality_report", fake_render)

    def boom(*a, **kw):
        raise AssertionError("不应发生二次分析")

    monkeypatch.setattr("offipy.audit.audit_pptx", boom)
    monkeypatch.setattr("offipy.art.analyze.build_scene", boom)
    monkeypatch.setattr("offipy.art.analyze.analyze_deck", boom)

    cli.main(["deck", "audit", "x.html"])
    assert len(calls) == 1


# ---------------------------------------------------------------- PPTX 流


def test_deck_audit_ppt_calls_analyze_deck_once(monkeypatch, capsys):
    from offipy import cli

    calls = []

    def fake_analyze(**kw):
        calls.append(kw)
        return _report()

    monkeypatch.setattr("offipy.art.analyze.analyze_deck", fake_analyze)
    cli.main(["deck", "audit", "--pptx", "x.pptx"])
    assert len(calls) == 1
    assert calls[0]["pptx"] == "x.pptx"
    assert calls[0]["profile"] == "balanced"
    out = capsys.readouterr().out
    assert "offipy deck audit" in out


def test_deck_audit_ppt_missing_file_friendly(monkeypatch, capsys):
    from offipy import cli

    def boom(**kw):
        raise FileNotFoundError("x.pptx")

    monkeypatch.setattr("offipy.art.analyze.analyze_deck", boom)
    assert cli.main(["deck", "audit", "--pptx", "x.pptx"]) == 2
    err = capsys.readouterr().err
    assert "找不到文件" in err
    assert "Traceback" not in err


def test_deck_audit_invalid_profile_friendly(monkeypatch, capsys):
    # 未知 --profile → get_profile 抛 KeyError，CLI 必须转友好 offipy: error + exit 2，
    # 绝不留裸 traceback（预运行无效输入 → exit 2，#39）
    from offipy import cli

    # PPTX 流
    def boom_ppt(**kw):
        raise KeyError("unknown art profile: bogus")

    monkeypatch.setattr("offipy.art.analyze.analyze_deck", boom_ppt)
    assert cli.main(["deck", "audit", "--pptx", "x.pptx", "--profile", "bogus"]) == 2
    err = capsys.readouterr().err
    assert "offipy: error:" in err
    assert "bogus" in err
    assert "Traceback" not in err

    # HTML 流
    def boom_html(*a, **kw):
        raise KeyError("unknown art profile: bogus")

    monkeypatch.setattr("offipy.deck.render_with_quality_report", boom_html)
    assert cli.main(["deck", "audit", "x.html", "--profile", "bogus"]) == 2
    err2 = capsys.readouterr().err
    assert "offipy: error:" in err2
    assert "bogus" in err2
    assert "Traceback" not in err2


# ---------------------------------------------------------------- 输出确定性


def test_deck_audit_json_deterministic(monkeypatch, capsys):
    from offipy import cli

    monkeypatch.setattr(
        "offipy.deck.render_with_quality_report", lambda html, **kw: _render_result()
    )
    cli.main(["deck", "audit", "x.html", "--json"])
    first = capsys.readouterr().out
    data = json.loads(first)
    assert data["source"] == "x.html"
    assert data["profile"] == "balanced"
    assert "warnings" in data
    assert "suggestions" in data
    # 页内 finding + deck 级 finding 都在
    assert len(data["suggestions"]) == 2
    for rec in data["suggestions"]:
        assert set(rec.keys()) == {
            "dimension",
            "slide_index",
            "rule_id",
            "element",
            "severity",
            "message",
            "suggestion",
        }
    # 顺序稳定：slide 先、deck 后
    assert data["suggestions"][0]["rule_id"] == "art.hierarchy.title_size"
    assert data["suggestions"][1]["rule_id"] == "art.consistency.suite"

    cli.main(["deck", "audit", "x.html", "--json"])
    second = capsys.readouterr().out
    assert first == second


def test_deck_audit_text_deterministic(monkeypatch, capsys):
    # 文本路径同样确定性：两次运行 stdout 逐字节一致
    from offipy import cli

    monkeypatch.setattr(
        "offipy.deck.render_with_quality_report", lambda html, **kw: _render_result()
    )
    cli.main(["deck", "audit", "x.html"])
    first = capsys.readouterr().out
    cli.main(["deck", "audit", "x.html"])
    second = capsys.readouterr().out
    assert first == second


def test_deck_audit_deck_finding_text_renders_all_deck(monkeypatch, capsys):
    # deck 级 finding（slide_index=None）文本渲染为（全篇），不显示 页 None；
    # JSON 记录仍保留 slide_index=null（数据不美化，仅文本层美化）
    from offipy import cli

    monkeypatch.setattr(
        "offipy.deck.render_with_quality_report", lambda html, **kw: _render_result()
    )
    cli.main(["deck", "audit", "x.html"])
    out = capsys.readouterr().out
    assert "页 None" not in out
    assert "（全篇）" in out
    assert "art.consistency.suite" in out


# ---------------------------------------------------------------- 建议投影


def test_deck_audit_suggestion_default():
    from offipy.art.suggest import project_suggestions

    # 无 remediation cue → 精确 "无自动建议"
    recs = project_suggestions(_report(), source="x.html")
    slide_rec = next(r for r in recs if r["rule_id"] == "art.hierarchy.title_size")
    assert slide_rec["suggestion"] == "无自动建议"

    # details 携带非空 "suggestion" 键 → 透出
    rep = _report()
    finding = rep.art.slides[0].dimensions[0].findings[0]
    finding.details["suggestion"] = "把标题字号提高到 32pt"
    recs2 = project_suggestions(rep, source="x.html")
    assert recs2[0]["suggestion"] == "把标题字号提高到 32pt"

    # 非字符串 / 空串 remediation → 视为无，仍回退 "无自动建议"
    rep3 = _report()
    f3 = rep3.art.slides[0].dimensions[0].findings[0]
    f3.details["fix"] = {"increase": 32}
    f3.details["remediation"] = "   "
    recs3 = project_suggestions(rep3, source="x.html")
    assert recs3[0]["suggestion"] == "无自动建议"

    # art 为空 → 空列表（CLI 文本流会打印（无建议记录））
    assert project_suggestions(_report(with_art=False), source="x.html") == []


# ---------------------------------------------------------------- 临时清理 / chromium 错误


def test_deck_audit_temp_workspace_cleaned(monkeypatch):
    from offipy import cli

    def boom(*a, **kw):
        raise ConversionError("渲染失败（模拟）")

    monkeypatch.setattr("offipy.deck.render_with_quality_report", boom)
    tmpdir = Path(tempfile.gettempdir())
    before = {d.name for d in tmpdir.iterdir() if d.name.startswith("offipy-deck-audit-")}
    assert cli.main(["deck", "audit", "x.html"]) == 1
    after = {d.name for d in tmpdir.iterdir() if d.name.startswith("offipy-deck-audit-")}
    assert after == before  # 失败路径也不残留临时工作区


def test_deck_audit_success_cleans_temp(monkeypatch, capsys):
    # 成功路径同样清理临时工作区（TemporaryDirectory 正常退出即删）
    from offipy import cli

    monkeypatch.setattr(
        "offipy.deck.render_with_quality_report", lambda html, **kw: _render_result()
    )
    tmpdir = Path(tempfile.gettempdir())
    before = {d.name for d in tmpdir.iterdir() if d.name.startswith("offipy-deck-audit-")}
    cli.main(["deck", "audit", "x.html", "--json"])
    after = {d.name for d in tmpdir.iterdir() if d.name.startswith("offipy-deck-audit-")}
    assert after == before  # 成功路径不残留临时工作区


def test_deck_audit_chromium_error_concise(monkeypatch, capsys):
    from offipy import cli

    def boom(*a, **kw):
        raise ConversionError(
            "HTML→PPTX 渲染需要 Chromium：...。请运行: python -m playwright install chromium"
        )

    monkeypatch.setattr("offipy.deck.render_with_quality_report", boom)
    assert cli.main(["deck", "audit", "x.html"]) == 1
    err = capsys.readouterr().err
    assert "playwright install chromium" in err
    assert "几何" not in err  # 不提供假的 geometry-only 兜底
