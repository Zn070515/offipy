from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"
# Task 6 会创建 examples/decks/starter/deck.html 并换回这里；
# 当前先顶替 vendored 工具的回归 fixture。
STARTER = (
    ROOT / "third_party" / "html-to-editable-pptx" / "tests" / "fixtures" / "regression_deck.html"
)

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
