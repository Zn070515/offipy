# tests/test_outline.py
"""outline 内容工作流测试（纯 Python，不依赖 Office）。"""

import json

import pytest

from offipy.outline import parse_outline

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
