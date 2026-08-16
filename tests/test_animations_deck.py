"""deck.py 动画接入：参数穿透 / no_visual_audit 条件拒绝 / 告警码。"""

import json

import pytest

from offipy.deck import _measure_warnings, _reject_no_visual_audit_declarations
from offipy.exceptions import InvalidArgumentError

MARKER_HTML = """<!doctype html><html><body>
<section data-pptx-slide>
  <div data-ppt-anim="fade">x</div>
</section></body></html>"""

TRANSITION_HTML = """<!doctype html><html><body>
<section data-pptx-slide data-ppt-transition="push">x</section></body></html>"""

FALLBACK_HTML = """<!doctype html><html><body>
<section data-pptx-slide><div data-aos="fade-up">x</div></section></body></html>"""

PLAIN_HTML = """<!doctype html><html><body>
<section data-pptx-slide><p>plain</p></section></body></html>"""


def test_reject_no_visual_audit_with_animations():
    with pytest.raises(InvalidArgumentError):
        _reject_no_visual_audit_declarations(MARKER_HTML, include_animations=True)


def test_reject_no_visual_audit_with_transition_marker():
    with pytest.raises(InvalidArgumentError):
        _reject_no_visual_audit_declarations(TRANSITION_HTML, include_animations=True)


def test_reject_no_visual_audit_with_fallback_marker():
    with pytest.raises(InvalidArgumentError):
        _reject_no_visual_audit_declarations(FALLBACK_HTML, include_animations=True)


def test_no_reject_without_animations_flag():
    _reject_no_visual_audit_declarations(MARKER_HTML)  # include_animations 默认 False


def test_no_reject_plain_html():
    _reject_no_visual_audit_declarations(PLAIN_HTML, include_animations=True)


def test_measure_warnings_anim_code():
    import pathlib
    import tempfile

    p = pathlib.Path(tempfile.mkdtemp()) / "measurements.json"
    p.write_text(json.dumps({"_warnings": [{"kind": "anim", "message": "skip"}]}), encoding="utf-8")
    ws = _measure_warnings(p)
    assert any(w.code == "deck.animation.skipped" for w in ws)
