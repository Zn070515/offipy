"""ppt.export_slides 纯逻辑（无 COM）：overwrite 拒绝、staging 原子替换、失败清理。"""

from pathlib import Path

import pytest

from offipy import ppt
from offipy.exceptions import FileConflictError


class _Slide:
    def __init__(self, i):
        self.i = i

    def Export(self, tmp, fmt, width, height):
        Path(tmp).write_bytes(b"png")


class _Slides:
    def __init__(self, n, fail_at=None):
        self.Count = n
        self._fail_at = fail_at

    def __call__(self, i):
        if self._fail_at is not None and i == self._fail_at:
            raise RuntimeError("export boom")
        return _Slide(i)


class _Pres:
    def __init__(self, n=2, fail_at=None):
        self.Slides = _Slides(n, fail_at)


def _app(pres):
    a = ppt.PptApp.__new__(ppt.PptApp)
    a._require_pres = lambda doc_id=None: pres
    return a


def _staging_dirs(tmp_path):
    return sorted(p.name for p in tmp_path.glob(".offipy-slides-*") if p.is_dir())


def test_export_slides_atomic_replace_cleans_staging(tmp_path):
    out = tmp_path / "png"
    app = _app(_Pres())
    paths = app.export_slides(str(out), width=100, height=50, doc_id="pres1")
    assert paths == [str(out / "slide_01.png"), str(out / "slide_02.png")]
    for p in paths:
        assert Path(p).read_bytes() == b"png"
    assert _staging_dirs(tmp_path) == []  # 成功后 staging 已清理


def test_export_slides_existing_refuses_fail_fast(tmp_path):
    out = tmp_path / "png"
    out.mkdir()
    (out / "slide_01.png").write_bytes(b"old")
    app = _app(_Pres())
    with pytest.raises(FileConflictError):
        app.export_slides(str(out), doc_id="pres1")
    assert (out / "slide_01.png").read_bytes() == b"old"  # 未触碰已有文件
    assert _staging_dirs(tmp_path) == []


def test_export_slides_failure_cleans_staging(tmp_path):
    out = tmp_path / "png"
    app = _app(_Pres(fail_at=2))
    with pytest.raises(RuntimeError):
        app.export_slides(str(out), doc_id="pres1")
    assert not (out / "slide_01.png").exists()  # 半成品未落最终位置
    assert _staging_dirs(tmp_path) == []  # 失败也清理 staging
