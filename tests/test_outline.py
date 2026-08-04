# tests/test_outline.py
"""outline 内容工作流测试（纯 Python，不依赖 Office）。"""

import json

import pytest

from offipy.outline import parse_outline, to_deck_html

SAMPLE = """# 季度业绩回顾
> 2026 Q2 财务摘要

## 营收增长 @layout: big-number
- 本季度营收 +18%
- 环比提升 5 个百分点

## 成本控制
一段正文说明降本举措。
- 采购流程优化
- 供应商集中
@notes: 补充：Q3 展望
"""


def test_parse_basic_structure():
    o = parse_outline(SAMPLE)
    assert o.title == "季度业绩回顾"
    assert o.subtitle == "2026 Q2 财务摘要"
    assert [s.index for s in o.slides] == [1, 2]
    assert o.slides[0].title == "营收增长"


def test_parse_bullets_and_body():
    o = parse_outline(SAMPLE)
    assert o.slides[0].bullets == ["本季度营收 +18%", "环比提升 5 个百分点"]
    assert o.slides[1].body == ["一段正文说明降本举措。"]
    assert o.slides[1].bullets == ["采购流程优化", "供应商集中"]


def test_parse_layout_and_note_directives():
    o = parse_outline(SAMPLE)
    assert o.slides[0].layout == "big-number"
    assert o.slides[1].note == "补充：Q3 展望"


def test_parse_kicker_directive():
    o = parse_outline("# T\n\n## 页\n@kicker: 眉标\n- 甲\n")
    assert o.slides[0].kicker == "眉标"


def test_parse_missing_title_raises():
    with pytest.raises(ValueError):
        parse_outline("## 只有页面\n- 无标题")


def test_to_json():
    o = parse_outline(SAMPLE)
    data = json.loads(o.to_json())
    assert data["title"] == "季度业绩回顾"
    assert data["slides"][0]["bullets"][0] == "本季度营收 +18%"


def test_markdown_roundtrip():
    o = parse_outline(SAMPLE)
    o2 = parse_outline(o.markdown())
    assert o.to_dict() == o2.to_dict()


def test_unknown_directive_raises():
    with pytest.raises(ValueError):
        parse_outline("# T\n\n## 页\n@nots: 拼错了\n")


def test_star_bullets_and_pending_directive():
    o = parse_outline("# T\n@layout: cards-3\n\n## 页\n* 甲\n* 乙\n")
    assert o.slides[0].layout == "cards-3"
    assert o.slides[0].bullets == ["甲", "乙"]


def test_quote_inside_slide_is_body():
    o = parse_outline("# T\n\n## 页\n- 甲\n> 补充说明\n")
    assert o.slides[0].bullets == ["甲"]
    assert o.slides[0].body == ["补充说明"]


def test_subtitle_only_before_first_slide():
    o = parse_outline("# T\n> 副标题\n\n## 页一\n> 页内引用\n")
    assert o.subtitle == "副标题"
    assert o.slides[0].body == ["页内引用"]


def test_to_deck_html_skeleton_structure():
    html = to_deck_html(parse_outline(SAMPLE))
    assert html.count("data-pptx-slide") == 2
    assert 'data-layout="big-number"' in html
    assert "季度业绩回顾" not in html  # deck 标题不进页面 HTML
    assert "营收增长" in html
    assert "<!DOCTYPE html>" in html
    assert '<div class="cards">' in html  # bullets 包 .cards 容器，与 cards-3 模板 DOM 对齐


def test_to_deck_html_autopick_default_layout():
    md = "# T\n\n## 开篇\n一句话。\n\n## 要点页\n- 甲\n- 乙\n- 丙"
    html = to_deck_html(parse_outline(md))
    assert 'data-layout="cards-3"' in html


def test_to_deck_html_injects_theme():
    html = to_deck_html(parse_outline(SAMPLE), theme="mckinsey")
    assert '<style data-theme="mckinsey">' in html
    assert "--bg: #FFFFFF" in html


def test_to_deck_html_escapes_text():
    o = parse_outline("# T\n\n## 页\n- a & b <c>\n")
    html = to_deck_html(o)
    assert "a &amp; b &lt;c&gt;" in html


def test_invalid_layout_directive_raises():
    # @layout 值拼进 class/data-layout 属性，必须拒绝注入尝试（引号/空格逃逸）
    with pytest.raises(ValueError):
        parse_outline('# T\n\n## 页\n@layout: cards-3" onmouseover="evil\n')
