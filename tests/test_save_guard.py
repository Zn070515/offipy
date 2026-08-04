"""save/save_pdf 覆盖保护（P1 资源）：目标已存在且不 overwrite → FileExistsError。

guard 必须在触 COM 之前触发（fail-fast）；overwrite=True 放行。用
`__new__` 构造实例跳过 __init__ 的 COM 初始化，专注测保护层本身。
"""

import os

import pytest

from offipy import excel, paths, ppt, word


def test_ensure_writable_refuses_existing(tmp_path):
    f = tmp_path / "x.pptx"
    f.write_text("x")
    with pytest.raises(FileExistsError):
        paths.ensure_writable(str(f))


def test_ensure_writable_allows_overwrite(tmp_path):
    f = tmp_path / "x.pptx"
    f.write_text("x")
    assert paths.ensure_writable(str(f), overwrite=True) == os.path.abspath(str(f))


def test_ensure_writable_missing_ok(tmp_path):
    f = tmp_path / "new.pptx"
    assert paths.ensure_writable(str(f)) == os.path.abspath(str(f))


class _FakePres:
    def __init__(self):
        self.saved = None

    def SaveAs(self, path, *a, **k):
        self.saved = path


def test_ppt_save_guard_fires_before_com(tmp_path):
    target = tmp_path / "deck.pptx"
    target.write_text("x")
    p = ppt.PptApp.__new__(ppt.PptApp)
    with pytest.raises(FileExistsError):
        p.save(str(target))


def test_ppt_save_overwrite_proceeds(tmp_path):
    target = tmp_path / "deck.pptx"
    target.write_text("x")
    p = ppt.PptApp.__new__(ppt.PptApp)
    fake = _FakePres()
    p.active_pres = lambda: fake
    p.save(str(target), overwrite=True)
    assert fake.saved == os.path.abspath(str(target))


def test_excel_save_pdf_guard_fires(tmp_path):
    target = tmp_path / "x.pdf"
    target.write_text("x")
    b = excel.ExcelApp.__new__(excel.ExcelApp)
    with pytest.raises(FileExistsError):
        b.save_pdf(str(target))


def test_word_save_pdf_guard_fires(tmp_path):
    target = tmp_path / "x.pdf"
    target.write_text("x")
    d = word.WordApp.__new__(word.WordApp)
    with pytest.raises(FileExistsError):
        d.save_pdf(str(target))
