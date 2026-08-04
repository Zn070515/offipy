from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"
STARTER = ROOT / "examples" / "decks" / "starter" / "deck.html"

pytestmark = pytest.mark.skipif(
    not CONVERT_PY.exists(),
    reason="third_party/html-to-editable-pptx 未 vendor",
)


def test_render_produces_pptx(tmp_path):
    from office_kit.deck import render

    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    assert out.exists()
    assert out.stat().st_size > 0


def test_starter_deck_slide_count(tmp_path):
    from office_kit.deck import render

    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    from pptx import Presentation

    assert len(Presentation(str(out)).slides) == 5
