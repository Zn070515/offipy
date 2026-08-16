from dataclasses import replace

from art_helpers import make_slide, make_text_element
from offipy.art.color import (
    RULES,
    accent_flood_rule,
    contrast_ratio,
    low_contrast_rule,
    no_accent_rule,
)
from offipy.art.features import compute_features
from offipy.art.models import ArtColor, ArtTextRun, ElementPixelEvidence
from offipy.art.profiles import get_profile
from offipy.art.rules import RuleContext, RuleSpec
from offipy.audit import Severity


def _ctx(slide, profile="balanced"):
    return RuleContext(
        profile=get_profile(profile),
        slide=slide,
        slide_index=slide.index,
        features=compute_features(slide),
        deck=__import__("offipy.art.models", fromlist=["ArtScene"]).ArtScene(slides=[slide]),
    )


def test_contrast_ratio_black_on_white():
    c = contrast_ratio(ArtColor(0, 0, 0), ArtColor(255, 255, 255))
    assert c > 20.0


def test_contrast_ratio_same():
    c = contrast_ratio(ArtColor(120, 120, 120), ArtColor(120, 120, 120))
    assert c == 1.0


def test_low_contrast_fires():
    slide = make_slide(
        1,
        elements=[
            make_text_element(
                "t", "Gray on white", font_size=24.0, foreground=ArtColor(200, 200, 200)
            ),
        ],
        background_color=ArtColor(255, 255, 255),
    )
    ev = low_contrast_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.color.low_contrast"


def test_low_contrast_uses_run_color():
    slide = make_slide(
        1,
        elements=[
            make_text_element(
                "t",
                "Mixed",
                font_size=24.0,
                foreground=ArtColor(0, 0, 0),
                runs=[
                    ArtTextRun(
                        text="Hi",
                        font_size=24.0,
                        font_size_unit="px",
                        color=ArtColor(230, 230, 230),
                    )
                ],
            ),
        ],
        background_color=ArtColor(255, 255, 255),
    )
    ev = low_contrast_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1


def test_low_contrast_unknown_background_reduces_coverage():
    # 无背景证据（页面无 background_color、元素无 background）→ covered 降
    slide = make_slide(
        1,
        elements=[
            make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(30, 30, 30)),
        ],
        background_color=None,
    )
    ev = low_contrast_rule(slide, _ctx(slide))
    assert ev.eligible_count == 1
    assert ev.covered_count == 0  # 背景未知 → 无对比可判


def test_low_contrast_foreground_vs_background_not_own_bg():
    # 黑字白底元素：foreground 黑、background 白 → 高对比，不误报
    slide = make_slide(
        1,
        elements=[
            make_text_element(
                "t",
                "Black on white",
                font_size=24.0,
                foreground=ArtColor(0, 0, 0),
                background=ArtColor(255, 255, 255),
            ),
        ],
        background_color=ArtColor(255, 255, 255),
    )
    assert low_contrast_rule(slide, _ctx(slide)).findings == []


def test_accent_flood_fires_over_ratio():
    slide = make_slide(
        1,
        elements=[
            make_text_element("a", "A", font_size=20.0, foreground=ArtColor(230, 0, 0)),
            make_text_element("b", "B", font_size=20.0, foreground=ArtColor(230, 0, 0), w=0.6),
        ],
    )
    ev = accent_flood_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].confidence <= 0.4  # experimental
    assert ev.findings[0].severity == Severity.LOW


def test_no_accent_fires_at_zero():
    slide = make_slide(
        1,
        elements=[
            make_text_element("a", "A", font_size=20.0, foreground=ArtColor(30, 30, 30)),
            make_text_element("b", "B", font_size=20.0, foreground=ArtColor(60, 60, 60)),
        ],
    )
    ev = no_accent_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].rule_id == "art.color.no_accent"
    assert ev.findings[0].confidence <= 0.4


def test_color_rules_are_rule_specs():
    assert all(isinstance(rs, RuleSpec) for rs in RULES)
    assert {rs.rule_id for rs in RULES} == {
        "art.color.low_contrast",
        "art.color.accent_flood",
        "art.color.no_accent",
    }


def test_low_contrast_declared_bg_used_over_pixel_bg():
    # 声明背景（黑）优先于像素背景（白）：灰字黑底高对比 → 无 finding。
    # 修复前像素白底会覆盖声明背景并误报低对比（#38）。
    el = make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(200, 200, 200))
    el = replace(
        el,
        pixel_evidence=ElementPixelEvidence(
            foreground=ArtColor(200, 200, 200),
            background=ArtColor(255, 255, 255),
            color_confidence=0.85,
            method="declared_verified",
        ),
    )
    slide = make_slide(1, elements=[el], background_color=ArtColor(0, 0, 0))
    ev = low_contrast_rule(slide, _ctx(slide))
    assert ev.findings == []


def test_low_contrast_declared_not_found_hint():
    el = make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(0, 0, 0))
    el = replace(
        el,
        pixel_evidence=ElementPixelEvidence(
            foreground=ArtColor(0, 0, 0),
            foreground_match_ratio=0.05,
            color_confidence=0.3,
            method="declared_not_found",
        ),
    )
    slide = make_slide(1, elements=[el], background_color=ArtColor(255, 255, 255))
    ev = low_contrast_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    f = ev.findings[0]
    assert f.confidence == 0.25  # 低置信提示，不驱动降级
    assert f.evidence_sources == frozenset({"pixel"})


def test_low_contrast_low_pixel_confidence_falls_back():
    el = make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(200, 200, 200))
    el = replace(
        el,
        pixel_evidence=ElementPixelEvidence(
            foreground=ArtColor(200, 200, 200),
            background=ArtColor(255, 255, 255),
            color_confidence=0.4,  # 低于 0.6 → 回退声明路径
            method="complex_background",
        ),
    )
    slide = make_slide(1, elements=[el], background_color=ArtColor(255, 255, 255))
    ev = low_contrast_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    assert ev.findings[0].evidence_sources == frozenset()  # 声明路径无 pixel 证据


def test_low_contrast_declared_bg_low_contrast_no_pixel_evidence():
    # 声明背景（白）+ 灰字 → 低对比 finding，证据走声明路径：无 pixel 证据、无「像素验证」后缀
    el = make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(200, 200, 200))
    el = replace(
        el,
        pixel_evidence=ElementPixelEvidence(
            foreground=ArtColor(200, 200, 200),
            background=ArtColor(255, 255, 255),
            color_confidence=0.85,
            method="declared_verified",
        ),
    )
    slide = make_slide(1, elements=[el], background_color=ArtColor(255, 255, 255))
    ev = low_contrast_rule(slide, _ctx(slide))
    assert len(ev.findings) == 1
    f = ev.findings[0]
    assert f.evidence_sources == frozenset()
    assert "像素验证" not in f.message


def test_low_contrast_obscured_by_image_no_false_positive():
    # #38 回归：白色文本被不透明图片覆盖 → 区域像素主色是图片深海军蓝。
    # 修复前 _text_evidence 把主色归因为文本背景 → 白字 vs 深色 → HIGH「对比度 1.14」误报。
    # 修复后像素背景不参与对比度：声明/有效背景未知 → 无对比可判 → 无 finding。
    el = make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(255, 255, 255))
    el = replace(
        el,
        pixel_evidence=ElementPixelEvidence(
            foreground=ArtColor(255, 255, 255),
            background=ArtColor(10, 30, 80),  # 旧 bug 输出：图片主色被归因为文本背景
            foreground_match_ratio=1.0,
            color_confidence=0.85,
            method="declared_verified",
        ),
    )
    slide = make_slide(1, elements=[el], background_color=None)
    ev = low_contrast_rule(slide, _ctx(slide))
    assert ev.covered_count == 0
    assert ev.findings == []


def test_low_contrast_declared_not_found_no_match_ratio_falls_back():
    # declared_not_found 但 foreground_match_ratio=None → 不发低置信提示，回退声明路径
    el = make_text_element("t", "Text", font_size=24.0, foreground=ArtColor(200, 200, 200))
    el = replace(
        el,
        pixel_evidence=ElementPixelEvidence(
            foreground=ArtColor(200, 200, 200),
            foreground_match_ratio=None,
            color_confidence=0.3,
            method="declared_not_found",
        ),
    )
    slide = make_slide(1, elements=[el], background_color=ArtColor(255, 255, 255))
    ev = low_contrast_rule(slide, _ctx(slide))
    # 未走 declared_not_found 提示分支 → 无 0.25 置信提示
    assert not any(f.confidence == 0.25 for f in ev.findings)
    # 回退声明路径：灰字白底 → 低对比 finding（无 pixel 证据）
    assert len(ev.findings) == 1
    assert ev.findings[0].evidence_sources == frozenset()


def test_accent_rule_eval_excludes_skip_roles():
    # 口径统一：page_number/footer 不在强调色评估范围（与调色板 _SKIP_ROLES 一致）
    slide = make_slide(
        1,
        elements=[
            make_text_element("b", "B", font_size=20.0, foreground=ArtColor(30, 30, 30)),
            make_text_element(
                "pn", "1", font_size=20.0, role="page_number", foreground=ArtColor(30, 30, 30)
            ),
        ],
    )
    ev = no_accent_rule(slide, _ctx(slide))
    assert ev.eligible_count == 1
    assert ev.covered_count == 1
