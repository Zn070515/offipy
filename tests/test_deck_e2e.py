from pathlib import Path

import pytest

from offipy.client import _ping
from offipy.deck import CONVERT_PY

STARTER = Path(__file__).resolve().parent.parent / "examples" / "decks" / "starter" / "deck.html"

pytestmark = pytest.mark.skipif(
    not CONVERT_PY.exists() or not _ping(),
    reason="需要 vendored 转换器 + 存活 server(8890)",
)


def test_full_loop_render_open_export(tmp_path):
    from offipy.deck import make

    png_dir = str(tmp_path / "png")
    pptx = make(str(STARTER), out=str(tmp_path / "deck.pptx"), feedback_dir=png_dir)
    assert Path(pptx).exists()
    pngs = sorted((tmp_path / "png").glob("slide_*.png"))
    assert len(pngs) == 5
    for p in pngs:
        assert p.stat().st_size > 0
