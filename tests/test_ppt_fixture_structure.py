"""P2-2：固定 fixture 结构验证（minimal_text_shapes.pptx）。

只验证 fixture 的预期结构（占位符页/纯文本框页/group 嵌套/页码脚日期占位符/
图片），不要求字节级重建。CI（Linux 纯模块 job）无 python-pptx 时整体跳过；
生成脚本 tests/fixtures/ppt/generate_minimal_text_shapes.py 是 Windows 本机
维护工具，CI 不跑。
"""

from pathlib import Path

import pytest

pptx = pytest.importorskip("pptx")

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "ppt" / "minimal_text_shapes.pptx"

# python-pptx 的 MsoShapeType / PpPlaceholderType 枚举值（与 COM 数值一致）
PICTURE = 13
TEXT_BOX = 17
GROUP = 6
PLACEHOLDER = 14


def _tree(shapes):
    """展平 group 树：[(shape, depth)]，含 group 内子元素。"""
    out = []

    def walk(shapes, depth):
        for sh in shapes:
            out.append((sh, depth))
            if hasattr(sh, "shapes"):  # GroupShape：python-pptx 用 .shapes 而非 COM GroupItems
                walk(sh.shapes, depth + 1)

    walk(shapes, 0)
    return out


def _ph_type(sh):
    return int(sh.placeholder_format.type)


def test_fixture_has_five_slides_and_page_size():
    prs = pptx.Presentation(str(FIXTURE))
    assert len(prs.slides) == 5
    assert prs.slide_width == 914400 * 10  # 10in EMU
    assert prs.slide_height == 914400 * 7.5  # 7.5in


def test_fixture_slide1_title_body_placeholders_and_notes():
    prs = pptx.Presentation(str(FIXTURE))
    s = prs.slides[0]
    placeholders = {_ph_type(sh) for sh in s.placeholders}
    assert placeholders == {1, 2}  # title + body
    texts = {sh.text_frame.text for sh in s.placeholders}
    assert "产品发布计划" in texts
    assert "第一季度\n第二季度\n第三季度" in texts
    assert s.notes_slide.notes_text_frame.text == "演讲者备注：本页为公司级发布计划"


def test_fixture_slide2_pure_textboxes_no_placeholders():
    prs = pptx.Presentation(str(FIXTURE))
    s = prs.slides[1]
    shapes = list(s.shapes)
    assert len(shapes) == 2
    assert all(sh.shape_type == TEXT_BOX for sh in shapes)
    assert all(not sh.is_placeholder for sh in shapes)
    assert {sh.name for sh in shapes} == {"TextBoxTitle", "TextBoxBody"}


def test_fixture_slide3_group_nesting():
    prs = pptx.Presentation(str(FIXTURE))
    s = prs.slides[2]
    flat = _tree(s.shapes)
    names = {sh.name for sh, _ in flat}
    assert names == {"GroupOuter", "GroupInner", "NestedTitle", "NestedBody", "OuterText"}
    by_name = {sh.name: sh for sh, _ in flat}
    # group 嵌套：GroupOuter 是 group，GroupInner 是其内层 group
    assert by_name["GroupOuter"].shape_type == GROUP
    assert by_name["GroupInner"].shape_type == GROUP
    # 文本子元素都是文本框
    assert all(
        by_name[n].shape_type == TEXT_BOX for n in ("NestedTitle", "NestedBody", "OuterText")
    )
    assert by_name["NestedTitle"].text_frame.text == "嵌套组内标题"
    assert by_name["NestedBody"].text_frame.text == "嵌套组内正文"
    assert by_name["OuterText"].text_frame.text == "外层组内文本"


def test_fixture_slide4_exempt_placeholders_and_page_number():
    prs = pptx.Presentation(str(FIXTURE))
    s = prs.slides[3]
    ph_types = {_ph_type(sh) for sh in s.placeholders}
    assert ph_types == {13, 15, 16}  # sldNum / footer / date
    # 角上纯数字文本框（非占位符）→ 页码候选豁免路径
    page_num = [sh for sh in s.shapes if sh.name == "TextBoxPageNum"]
    assert len(page_num) == 1
    assert not page_num[0].is_placeholder
    assert page_num[0].text_frame.text == "7"


def test_fixture_slide5_picture_and_caption():
    prs = pptx.Presentation(str(FIXTURE))
    s = prs.slides[4]
    shapes = list(s.shapes)
    assert len(shapes) == 2
    picture = [sh for sh in shapes if sh.shape_type == PICTURE]
    assert len(picture) == 1
    assert not picture[0].has_text_frame  # 无文本能力 → read_slide_texts 不返回
    caption = [sh for sh in shapes if sh.name == "PictureCaption"]
    assert len(caption) == 1
    assert caption[0].text_frame.text == "图片说明"
