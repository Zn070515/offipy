from pathlib import Path

import pytest

from offipy.deck import CONVERT_PY
from offipy.envcheck import _check_browser

STARTER = Path(__file__).resolve().parent.parent / "examples" / "decks" / "starter" / "deck.html"

pytestmark = [
    pytest.mark.deck_render,
    pytest.mark.skipif(
        not CONVERT_PY.exists() or not _check_browser().ok,
        reason="vendored 转换器缺失或 chromium 不可用",
    ),
]


def test_render_produces_pptx(tmp_path):
    from offipy.deck import render

    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    assert out.exists()
    assert out.stat().st_size > 0


def test_starter_deck_slide_count(tmp_path):
    pytest.importorskip("pptx")
    from offipy.deck import render

    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    from pptx import Presentation

    assert len(Presentation(str(out)).slides) == 5


def test_default_out_name(tmp_path):
    from offipy.deck import _default_out

    assert _default_out(str(tmp_path / "deck.html")) == str(tmp_path / "deck.pptx")
    assert _default_out(str(tmp_path / "deck.audited.html")) == str(tmp_path / "deck.pptx")


def test_render_unzip_slide1_contains_title_text(tmp_path):
    # review L592-603 e2e 契约：render 产出的 .pptx 解包后，slide1 的 XML 里真实
    # 落盘了源 HTML 的标题文本（不是空白/占位）。标题含 <br> 会拆成两个 a:t run，
    # 断言子串「产品增长报告」落在 slide1.xml 即可跨 run 边界成立。
    import zipfile

    from offipy.deck import render

    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    with zipfile.ZipFile(out) as z:
        slide1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "产品增长报告" in slide1
