from pathlib import Path

import pytest

from offipy.deck import CONVERT_PY

STARTER = Path(__file__).resolve().parent.parent / "examples" / "decks" / "starter" / "deck.html"

pytestmark = pytest.mark.skipif(
    not CONVERT_PY.exists(),
    reason="vendored 转换器缺失（src/offipy/_vendor/）",
)


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
