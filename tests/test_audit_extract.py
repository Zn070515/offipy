"""audit 提取：递归 shape 提取（group 嵌套）、隐藏、连接线、文本/autofit 读取。"""

import builtins

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt

from offipy.audit.extract import (
    _GroupTransform,
    _PresentationExtract,
    _TextRun,
    extract_presentation,
)


def _build_pptx() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    # 文本框：多 run + wrap/autofit
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    tf = tb.text_frame
    tf.text = "hello"
    r = tf.paragraphs[0].add_run()
    r.text = " world"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.name = "Arial"
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # normAutofit
    # 自动形状：有文字，无 run 级字号（继承）
    sh = slide.shapes.add_shape(1, Inches(5), Inches(5), Inches(2), Inches(1))
    sh.text = "shape"
    # 连接线
    slide.shapes.add_connector(1, Inches(0), Inches(5), Inches(2), Inches(6))
    # 表格
    slide.shapes.add_table(2, 2, Inches(6), Inches(5), Inches(3), Inches(1))
    return prs


def _make_extract(tmp_path, prs) -> _PresentationExtract:
    path = tmp_path / "x.pptx"
    prs.save(path)
    return extract_presentation(path)


# ---------------------------------------------------------------- 顶层 shape


def test_extract_textbox_text_and_runs(tmp_path):
    ext = _make_extract(tmp_path, _build_pptx())
    assert ext.slide_size[0] == pytest.approx(13.333, abs=1e-3)
    assert ext.slide_size[1] == pytest.approx(7.5)
    assert len(ext.slides) == 1
    assert ext.slides[0].slide_index == 1

    rec = next(s for s in ext.slides[0].shapes if s.shape_type == "TEXT_BOX")
    assert rec.text == "hello world"
    assert rec.has_text_frame is True
    assert rec.word_wrap is True
    assert rec.autofit_mode == "TEXT_TO_FIT_SHAPE"
    assert rec.rotation == 0.0
    assert rec.left == pytest.approx(1.0)
    assert rec.width == pytest.approx(3.0)
    # run 级字号/加粗/字体
    run: _TextRun = rec.paragraphs[0].runs[1]
    assert run.text == " world"
    assert run.font_size == pytest.approx(18.0)
    assert run.bold is True
    assert run.font_name == "Arial"
    # TextFrame 默认 margin 0.1 英寸
    assert rec.tf_margin_left == pytest.approx(0.1)


def test_extract_autoshape_inherited_font_size(tmp_path):
    ext = _make_extract(tmp_path, _build_pptx())
    rec = next(s for s in ext.slides[0].shapes if s.text == "shape")
    assert rec.text == "shape"
    assert rec.has_text_frame is True
    assert rec.paragraphs[0].runs[0].font_size is None  # 继承


# ---------------------------------------------------------------- group 嵌套


def test_extract_group_nesting_local_coords(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    g = slide.shapes.add_group_shape()
    c1 = g.shapes.add_shape(1, Inches(0.5), Inches(0.25), Inches(1), Inches(1))
    c1.text = "child"
    g2 = g.shapes.add_group_shape()  # 嵌套 group
    c2 = g2.shapes.add_textbox(Inches(0), Inches(0), Inches(1), Inches(0.5))
    c2.text = "deep"
    ext = _make_extract(tmp_path, prs)

    group_rec = next(s for s in ext.slides[0].shapes if s.is_group and s.shape_id == g.shape_id)
    assert group_rec.shape_type == "GROUP"
    assert group_rec.parent_shape_id is None
    assert group_rec.group_path == ()
    assert isinstance(group_rec.transform, _GroupTransform)
    assert group_rec.transform.ch_ext_cx == pytest.approx(1.5)  # 坐标空间随子元素扩展

    c1_rec = next(s for s in ext.slides[0].shapes if s.shape_id == c1.shape_id)
    assert c1_rec.parent_shape_id == g.shape_id
    assert c1_rec.group_path == (g.shape_id,)
    assert c1_rec.left == pytest.approx(0.5)  # 局部坐标
    assert c1_rec.top == pytest.approx(0.25)

    g2_rec = next(s for s in ext.slides[0].shapes if s.shape_id == g2.shape_id)
    assert g2_rec.group_path == (g.shape_id,)
    c2_rec = next(s for s in ext.slides[0].shapes if s.shape_id == c2.shape_id)
    assert c2_rec.group_path == (g.shape_id, g2.shape_id)
    assert c2_rec.text == "deep"


def test_group_children_z_order_scoped(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    g = slide.shapes.add_group_shape()
    a = g.shapes.add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    b = g.shapes.add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    ext = _make_extract(tmp_path, prs)
    ra = next(s for s in ext.slides[0].shapes if s.shape_id == a.shape_id)
    rb = next(s for s in ext.slides[0].shapes if s.shape_id == b.shape_id)
    assert ra.z_order == 0
    assert rb.z_order == 1


# ---------------------------------------------------------------- 隐藏


def test_hidden_shape_detection(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    h = slide.shapes.add_shape(1, Inches(1), Inches(1), Inches(1), Inches(1))
    h._element.xpath("./p:nvSpPr/p:cNvPr")[0].set("hidden", "1")
    v = slide.shapes.add_shape(1, Inches(2), Inches(2), Inches(1), Inches(1))
    ext = _make_extract(tmp_path, prs)
    hr = next(s for s in ext.slides[0].shapes if s.shape_id == h.shape_id)
    vr = next(s for s in ext.slides[0].shapes if s.shape_id == v.shape_id)
    assert hr.is_hidden is True
    assert vr.is_hidden is False


# ---------------------------------------------------------------- 连接线 / 表


def test_connector_detection(tmp_path):
    ext = _make_extract(tmp_path, _build_pptx())
    conn = next(s for s in ext.slides[0].shapes if s.is_connector)
    assert conn.shape_type == "LINE"
    assert conn.is_connector is True
    assert conn.is_group is False


def test_connector_type_enum_line():
    assert MSO_SHAPE_TYPE.LINE is not None  # 探针保证判定入口存在


def test_table_detection(tmp_path):
    ext = _make_extract(tmp_path, _build_pptx())
    tbl = next(s for s in ext.slides[0].shapes if s.shape_type == "TABLE")
    assert tbl.has_table is True
    assert tbl.has_text_frame is False
    assert tbl.text == ""


# ---------------------------------------------------------------- 占位符


def test_title_placeholder_type(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide
    title = slide.shapes.title
    title.text = "T"
    ext = _make_extract(tmp_path, prs)
    rec = next(s for s in ext.slides[0].shapes if s.shape_id == title.shape_id)
    assert rec.placeholder_type is not None
    assert rec.has_text_frame is True


# ---------------------------------------------------------------- 门禁


def test_import_extract_does_not_load_pptx(monkeypatch):
    orig_import = builtins.__import__

    def guarded(name, *args, **kwargs):
        if name == "pptx" or name.startswith("pptx."):
            raise AssertionError("import offipy.audit.extract 不应加载 python-pptx")
        return orig_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", guarded)
    import offipy.audit.extract as m  # noqa: F401

    assert m is not None


def test_lazy_pptx_only_on_read(tmp_path):
    # extract_presentation 内部才加载 pptx；此处直接调它应成功
    ext = _make_extract(tmp_path, _build_pptx())
    assert len(ext.slides[0].shapes) >= 4
