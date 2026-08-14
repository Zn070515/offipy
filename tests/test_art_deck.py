import json
from pathlib import Path

import pytest

import offipy.deck as deck
from offipy.art.models import ArtReport, ArtScene
from offipy.audit.models import AuditConfig, PptxAuditReport
from offipy.exceptions import ConversionError, InvalidArgumentError


def _fake_audit(p, config=None):
    return PptxAuditReport(
        schema_version="0.1",
        offipy_version="0.11.6",
        path=str(p),
        source_sha256="abc",
        slide_size=(10.0, 7.5),
        slide_count=1,
        config=config or AuditConfig(),
    )


def _fake_run_art_analysis(
    measurements, profile="balanced", pptx_report=None, slides_dir=None, pixel_required=False
):
    """_run_art_analysis 替身：接受 v0.12.1 新增的 pixel_required kwarg。"""
    return ArtReport(profile=profile)


class _FakeStage:
    """可调用的 _render_stage 替身：返回一个 CM 实例（不是 CM 本身不可调用）。"""

    def __init__(self, out_dir):
        self._out_dir = out_dir
        self.commit_called = False

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    @property
    def tmp_pptx(self):
        return f"{self._out_dir}/tmp.pptx"

    @property
    def final_pptx(self):
        return f"{self._out_dir}/deck.pptx"

    @property
    def measurements_path(self):
        m = (
            Path(self.tmp_pptx).parent
            / f"{Path(self.tmp_pptx).stem}_audit"
            / "_cache"
            / "measurements.json"
        )
        return m if m.is_file() else None

    def commit(self):
        self.commit_called = True

    def rollback(self):
        pass


def test_quality_render_result_subclasses_render_result(tmp_path, monkeypatch):
    from offipy.deck import QualityRenderResult, RenderResult

    assert issubclass(QualityRenderResult, RenderResult)


def test_render_with_quality_report_assembles_art(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    # defer_audit_preserve：with 块内审计目录是 tmp 名（tmp_audit），不是最终名
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    # render_with_quality_report 内部 `from .audit import audit_pptx`（惰性）→ 必须 patch 真实模块
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    monkeypatch.setattr(deck, "_run_art_analysis", _fake_run_art_analysis)
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    result = deck.render_with_quality_report("in.html", profile="balanced")
    assert result.art_report is not None
    assert result.art_report.profile == "balanced"
    assert result.deck_quality.art is result.art_report


def test_render_with_quality_report_missing_measurements(tmp_path, monkeypatch):
    out_dir = tmp_path / "out2"
    out_dir.mkdir(parents=True)
    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    result = deck.render_with_quality_report("in.html")
    assert result.art_report is None
    assert any(w.code == "art.measurements_missing" for w in result.deck_quality.warnings)


def test_run_art_analysis_dual_source(monkeypatch):
    captured = {}

    def fake_build(*, measurements=None, pptx_report=None, slides_dir=None):
        captured["measurements"] = measurements
        captured["pptx_report"] = pptx_report
        captured["slides_dir"] = slides_dir
        return ArtScene()

    # _run_art_analysis 从 offipy.art 包级 re-export 解析 build_scene → patch 包级名
    monkeypatch.setattr("offipy.art.build_scene", fake_build)
    art = deck._run_art_analysis(
        measurements={"slides": []}, profile="balanced", pptx_report=_fake_audit("x.pptx")
    )
    assert art is not None
    assert captured["pptx_report"] is not None  # 双源融合：pptx_report 确实传下去
    assert captured["slides_dir"] is None  # 默认像素关闭：slides_dir 不传


def test_render_with_quality_report_pixel_off_default(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    monkeypatch.setattr(deck, "_run_art_analysis", _fake_run_art_analysis)
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    result = deck.render_with_quality_report("in.html", profile="balanced")
    assert result.art_report is not None
    assert not list(out_dir.glob("offipy-pixel-*"))  # 默认不建 staging


def test_render_with_quality_report_pixel_required_failure_raises(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    monkeypatch.setattr(
        deck,
        "_export_pixel_slides",
        lambda pptx, out_dir: (_ for _ in ()).throw(RuntimeError("no COM")),
    )
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    with pytest.raises(ConversionError):
        deck.render_with_quality_report("in.html", pixel_analysis="required")
    assert not list(out_dir.glob("offipy-pixel-*"))  # staging 已清理


def test_render_with_quality_report_pixel_best_effort_failure_warns(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    monkeypatch.setattr(
        deck,
        "_export_pixel_slides",
        lambda pptx, out_dir: (_ for _ in ()).throw(RuntimeError("no COM")),
    )
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    result = deck.render_with_quality_report("in.html", pixel_analysis="best_effort")
    assert any(w.code == "art.pixel.best_effort_failed" for w in result.deck_quality.warnings)


def test_render_with_quality_report_pixel_preserve(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)

    def fake_export(pptx, out_dir):
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        Path(out_dir, "slide_1.png").write_bytes(b"png")
        return [str(Path(out_dir, "slide_1.png"))]

    monkeypatch.setattr(deck, "_export_pixel_slides", fake_export)
    monkeypatch.setattr(deck, "_write_deck_info", lambda out_dir, pptx: None)
    monkeypatch.setattr(deck, "_run_art_analysis", _fake_run_art_analysis)
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    deck.render_with_quality_report(
        "in.html", pixel_analysis="required", preserve_pixel_slides=True
    )
    final_slides = out_dir / "deck_slides"
    assert (final_slides / "slide_1.png").is_file()
    assert not list(out_dir.glob("offipy-pixel-*"))  # staging 清理


def test_render_with_quality_report_invalid_pixel_analysis(tmp_path, monkeypatch):
    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(tmp_path))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    with pytest.raises(InvalidArgumentError):
        deck.render_with_quality_report("in.html", pixel_analysis="always")


def test_render_with_quality_report_pixel_commit_failure_cleans_staging(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)

    def fake_export(pptx, out_dir):
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        Path(out_dir, "slide_1.png").write_bytes(b"png")
        return [str(Path(out_dir, "slide_1.png"))]

    monkeypatch.setattr(deck, "_export_pixel_slides", fake_export)
    monkeypatch.setattr(deck, "_write_deck_info", lambda out_dir, pptx: None)
    monkeypatch.setattr(deck, "_run_art_analysis", _fake_run_art_analysis)
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    def raise_commit(self):
        raise ConversionError("final pptx locked")

    monkeypatch.setattr(_FakeStage, "commit", raise_commit)

    with pytest.raises(ConversionError):
        deck.render_with_quality_report(
            "in.html", pixel_analysis="required", preserve_pixel_slides=True
        )
    assert not list(out_dir.glob("offipy-pixel-*"))  # commit 抛错也清理 staging


def test_move_slides_to_final_scrubs_stale(tmp_path):
    # 归属校验：#audit 分支A——只有 _deck_info.json 标记该目录归本 deck 才清旧页。
    final = tmp_path / "final"
    final.mkdir(parents=True)
    final_pptx = Path(str(final / "deck.pptx")).resolve()
    (final / "deck.pptx").write_bytes(b"pptx")  # _write_deck_info 会哈希该文件

    staging = tmp_path / "staging" / "slides"
    staging.mkdir(parents=True)
    (staging / "slide_1.png").write_bytes(b"fresh1")
    (staging / "_deck_info.json").write_text(
        json.dumps({"schema": 1, "pptx": str(final_pptx)}), encoding="utf-8"
    )

    (final / "slide_1.png").write_bytes(b"old1")
    (final / "slide_2.png").write_bytes(b"old2")
    (final / "_deck_info.json").write_text(
        json.dumps({"schema": 1, "pptx": str(final_pptx)}), encoding="utf-8"
    )
    (final / "notes.txt").write_text("keep", encoding="utf-8")

    class _DummyStage:
        def __init__(self):
            self.final_pptx = final_pptx

    deck._move_slides_to_final(str(staging), _DummyStage(), slides_output_dir=str(final))

    assert (final / "slide_1.png").read_bytes() == b"fresh1"
    # 标记被重写为当前 pptx 归属
    info = json.loads((final / "_deck_info.json").read_text(encoding="utf-8"))
    assert info["pptx"] == str(final_pptx)
    assert (final / "notes.txt").read_text(encoding="utf-8") == "keep"
    assert not (final / "slide_2.png").exists()  # 旧页残留被清掉


def test_render_with_quality_report_pixel_required_missing_measurements_raises(
    tmp_path, monkeypatch
):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    with pytest.raises(ConversionError):
        deck.render_with_quality_report("in.html", pixel_analysis="required")
    assert not list(out_dir.glob("offipy-pixel-*"))  # 无 measurements 不建 staging


def test_render_with_quality_report_pixel_required_empty_export_raises(tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)
    audit_dir = out_dir / "tmp_audit" / "_cache"
    audit_dir.mkdir(parents=True)
    (audit_dir / "measurements.json").write_text(json.dumps({"slides": []}), encoding="utf-8")

    monkeypatch.setattr(deck, "_render_stage", lambda *a, **kw: _FakeStage(out_dir))
    monkeypatch.setattr("offipy.audit.audit_pptx", _fake_audit)
    monkeypatch.setattr(deck, "_export_pixel_slides", lambda pptx, out_dir: [])
    monkeypatch.setattr(deck, "_write_deck_info", lambda out_dir, pptx: None)
    monkeypatch.setattr(deck, "_atomic_replace", lambda src, dst: None)

    with pytest.raises(ConversionError):
        deck.render_with_quality_report("in.html", pixel_analysis="required")
    assert not list(out_dir.glob("offipy-pixel-*"))  # staging 已清理


def test_run_art_analysis_required_no_pixel_raises(monkeypatch):
    monkeypatch.setattr("offipy.art.build_scene", lambda **kw: ArtScene())
    with pytest.raises(ConversionError):
        deck._run_art_analysis(
            {"slides": []}, "balanced", slides_dir="out/slides", pixel_required=True
        )


def test_run_art_analysis_required_with_pixel_source_ok(monkeypatch):
    monkeypatch.setattr("offipy.art.build_scene", lambda **kw: ArtScene(sources={"pixel"}))
    art = deck._run_art_analysis(
        {"slides": []}, "balanced", slides_dir="out/slides", pixel_required=True
    )
    assert isinstance(art, ArtReport)
