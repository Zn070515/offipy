import sys
from pathlib import Path

import pytest

from offipy import core
from offipy.client import call

pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(
        sys.platform != "win32" or not core.running("ppt"),
        reason="需要存活的 PowerPoint（server 8890 持有）",
    ),
]


def test_export_slides_pngs(tmp_path):
    did = call("ppt", "new_pres")
    call("ppt", "add_slide", layout=1, doc_id=did)  # title
    call("ppt", "set_title", slide_idx=1, text="Export Test", doc_id=did)
    call("ppt", "add_slide", layout=2, doc_id=did)  # title+body
    call("ppt", "set_title", slide_idx=2, text="Page Two", doc_id=did)
    out_dir = str(tmp_path / "png")
    paths = call("ppt", "export_slides", out_dir=out_dir, width=1920, height=1080, doc_id=did)
    assert len(paths) == 2
    for p in paths:
        assert Path(p).exists()
        assert Path(p).stat().st_size > 0


def _ppt():
    return core.connect("ppt")


def _shape_texts(pres, slide_idx):
    shapes = pres.Slides(slide_idx).Shapes
    return [shapes(i).TextFrame.TextRange.Text for i in range(1, shapes.Count + 1)]


def test_set_title_body_blank_layout_auto_adds_textbox():
    # PP_LAYOUT_BLANK(12) 无任何占位符：set_title/set_body 自动建文本框并返回 shape ID
    did = call("ppt", "new_pres")
    call("ppt", "add_slide", layout=12, doc_id=did)
    sid = call("ppt", "set_title", slide_idx=1, text="自动标题", doc_id=did)
    assert isinstance(sid, int)
    call("ppt", "add_slide", layout=12, doc_id=did)
    sid2 = call("ppt", "set_body", slide_idx=2, lines=["甲", "乙"], doc_id=did)
    assert isinstance(sid2, int)
    pres = _ppt().ActivePresentation
    assert "自动标题" in _shape_texts(pres, 1)
    joined2 = "\n".join(_shape_texts(pres, 2))
    assert "甲" in joined2 and "乙" in joined2
