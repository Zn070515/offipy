from pathlib import Path

import pytest

from offipy import core
from offipy.client import call

pytestmark = pytest.mark.skipif(
    not core.running("ppt"),
    reason="需要存活的 PowerPoint（server 8890 持有）",
)


def test_export_slides_pngs(tmp_path):
    call("ppt", "new_pres")
    call("ppt", "add_slide", layout=1)  # title
    call("ppt", "set_title", slide_idx=1, text="Export Test")
    call("ppt", "add_slide", layout=2)  # title+body
    call("ppt", "set_title", slide_idx=2, text="Page Two")
    out_dir = str(tmp_path / "png")
    paths = call("ppt", "export_slides", out_dir=out_dir, width=1920, height=1080)
    assert len(paths) == 2
    for p in paths:
        assert Path(p).exists()
        assert Path(p).stat().st_size > 0
