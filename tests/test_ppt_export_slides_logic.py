"""ppt.export_slides 纯逻辑（无 COM）：overwrite 拒绝、staging 原子替换、失败清理。"""

from pathlib import Path

import pytest

from offipy import ppt
from offipy.exceptions import FileConflictError, InvalidArgumentError


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


def test_export_slides_out_dir_is_file_refuses(tmp_path):
    # S5：out_dir 是已存在文件 → 拦成 FileConflictError（原 os.makedirs 抛裸
    # FileExistsError → server 归一 RemoteCallError 消息丑）。目录是文件时无论如何
    # 都不能导出进它，overwrite=True 也要拦（在 makedirs 之前拦截）。
    out_file = tmp_path / "afile"
    out_file.write_bytes(b"x")
    app = _app(_Pres())
    with pytest.raises(FileConflictError, match="输出目录已存在且不是目录"):
        app.export_slides(str(out_file), doc_id="pres1")
    with pytest.raises(FileConflictError, match="输出目录已存在且不是目录"):
        app.export_slides(str(out_file), overwrite=True, doc_id="pres1")


def test_export_slides_width_height_must_be_positive_ints(tmp_path):
    # C3：width/height 传 0/负数/非整数/超上限会让 COM Export 失败或分配巨幅位图
    app = _app(_Pres())
    for bad in (0, -1, 1.5, "100", True, 100_000_000):
        with pytest.raises(InvalidArgumentError):
            app.export_slides(str(tmp_path / "o1"), width=bad, height=50, doc_id="pres1")
        with pytest.raises(InvalidArgumentError):
            app.export_slides(str(tmp_path / "o2"), width=100, height=bad, doc_id="pres1")


def test_export_slides_width_height_large_but_valid_ok(tmp_path):
    app = _app(_Pres())
    paths = app.export_slides(str(tmp_path / "png"), width=10000, height=10000, doc_id="pres1")
    assert len(paths) == 2
