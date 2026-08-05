"""audit text-fit 规则：Pillow 字体度量 vs 字符权重回退、无可用空间、横/纵溢出。"""

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

from offipy.audit.extract import extract_presentation
from offipy.audit.models import (
    RULE_TEXT_FIT_HORIZONTAL,
    RULE_TEXT_FIT_VERTICAL,
    AuditConfig,
    Severity,
)
from offipy.audit.rules import run_rules


def _run(prs, tmp_path, config=None):
    path = tmp_path / "x.pptx"
    prs.save(path)
    ext = extract_presentation(path)
    records = [r for slide in ext.slides for r in slide.shapes]
    findings, suppressed = run_rules(records, ext.slide_size, config or AuditConfig())
    return findings, suppressed, records


def _add_tb(slide, x, y, text, w=1, h=0.5, font_name=None, font_size=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.text = text
    run = tb.text_frame.paragraphs[0].runs[0]
    if font_name is not None:
        run.font.name = font_name
    if font_size is not None:
        run.font.size = Pt(font_size)
    return tb


def _by_rule(findings, rule_id):
    return [f for f in findings if f.rule_id == rule_id]


def _any_font_path() -> str | None:
    for path in (
        r"C:\Windows\Fonts\arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ):
        if Path(path).exists():
            return path
    return None


def test_fallback_char_weight_low_confidence(tmp_path):
    # 找不到字体文件 → 字符权重回退，confidence 0.4，消息标注低置信
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_tb(slide, 1, 1, "WWWWWWWWWW", font_name="NonexistentFontXYZ", font_size=18)
    findings, _, _ = _run(prs, tmp_path)
    fits = _by_rule(findings, RULE_TEXT_FIT_HORIZONTAL)
    assert len(fits) == 1
    f = fits[0]
    assert f.severity == Severity.LOW
    assert f.confidence == 0.4
    assert "字符估算低置信" in f.message
    assert f.details["text_width_in"] == 1.25


def test_pillow_metrics_high_confidence(tmp_path):
    font_path = _any_font_path()
    if font_path is None:
        pytest.skip("无可用系统字体，跳过 Pillow 度量测试")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_tb(slide, 1, 1, "WWWWWWWWWWWWWWWWWWWW", font_name=font_path, font_size=18)
    findings, _, _ = _run(prs, tmp_path)
    fits = _by_rule(findings, RULE_TEXT_FIT_HORIZONTAL)
    assert len(fits) == 1
    f = fits[0]
    assert f.severity == Severity.LOW
    assert f.confidence == 0.8
    assert "字符估算低置信" not in f.message


def test_vertical_overflow_low(tmp_path):
    # 5 行文字在 0.5 英寸高框内 → 纵向溢出 LOW
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_tb(slide, 1, 1, "line1\nline2\nline3\nline4\nline5", w=2, h=0.5)
    findings, _, _ = _run(prs, tmp_path)
    verts = _by_rule(findings, RULE_TEXT_FIT_VERTICAL)
    assert len(verts) == 1
    f = verts[0]
    assert f.severity == Severity.LOW
    assert f.details["text_height_in"] > f.details["avail_height_in"]


def test_no_usable_space_mid(tmp_path):
    # 内边距吃尽可用区域 → 无可用空间 MID
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_tb(slide, 1, 1, "x", w=0.1, h=0.1)
    findings, _, _ = _run(prs, tmp_path)
    fits = _by_rule(findings, RULE_TEXT_FIT_HORIZONTAL)
    assert len(fits) == 1
    assert fits[0].severity == Severity.MID
