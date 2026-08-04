"""save/save_pdf 覆盖保护（P1 资源）：目标已存在且不 overwrite → FileExistsError。

guard 必须在触 COM 之前触发（fail-fast）；overwrite=True 放行。用
`__new__` 构造实例跳过 __init__ 的 COM 初始化，专注测保护层本身。
"""

import os
import types

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
    p.app = types.SimpleNamespace(DisplayAlerts=0)  # _alerts_scope 读写 DisplayAlerts
    fake = _FakePres()
    p.active_pres = lambda doc_id=None: fake
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


# --- P0-2：DisplayAlerts 常量正确 + 全局副作用还原 ---


def test_display_alerts_constants():
    # 回归：ppt 原 =0 实为 ppAlertsAll，ppAlertsNone 才是 1
    assert ppt.PP_ALERTS_NONE == 1
    assert ppt.PP_FIXED_FORMAT_TYPE_PDF == 2
    assert word.WD_ALERTS_NONE == 0
    assert word.WD_EXPORT_FORMAT_PDF == 17


def _ensure_app_returning(app):
    def _ensure(app_name, visible=True):
        return app, True

    return _ensure


# --- P0-5：DisplayAlerts 作用域化——构造不全局静音，作用域内临时抑制并还原 ---


def test_ppt_init_no_global_mute_and_quit_restores(monkeypatch):
    app = types.SimpleNamespace(DisplayAlerts=0)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_app_returning(app))
    monkeypatch.setattr("offipy.core.quit_app", lambda app_name: True)
    p = ppt.PptApp()
    assert app.DisplayAlerts == 0  # 构造不再全局静音（原值保留）
    p.quit()
    assert app.DisplayAlerts == 0  # quit 兜底还原到保存值


def test_ppt_alerts_scope_restores_on_exception(monkeypatch):
    app = types.SimpleNamespace(DisplayAlerts=0)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_app_returning(app))
    p = ppt.PptApp()
    with pytest.raises(RuntimeError), p._alerts_scope():
        assert app.DisplayAlerts == ppt.PP_ALERTS_NONE  # 作用域内抑制
        raise RuntimeError("boom")
    assert app.DisplayAlerts == 0  # 异常路径也还原


def test_word_init_no_global_mute_and_quit_restores(monkeypatch):
    app = types.SimpleNamespace(DisplayAlerts=1)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_app_returning(app))
    monkeypatch.setattr("offipy.core.quit_app", lambda app_name: True)
    d = word.WordApp()
    assert app.DisplayAlerts == 1  # 原值保留
    d.quit()
    assert app.DisplayAlerts == 1


def test_word_alerts_scope_restores_on_exception(monkeypatch):
    app = types.SimpleNamespace(DisplayAlerts=1)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_app_returning(app))
    d = word.WordApp()
    with pytest.raises(RuntimeError), d._alerts_scope():
        assert app.DisplayAlerts == word.WD_ALERTS_NONE
        raise RuntimeError("boom")
    assert app.DisplayAlerts == 1


def test_excel_init_no_global_mute_and_quit_restores(monkeypatch):
    app = types.SimpleNamespace(DisplayAlerts=True)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_app_returning(app))
    monkeypatch.setattr("offipy.core.quit_app", lambda app_name: True)
    b = excel.ExcelApp()
    assert app.DisplayAlerts is True  # 原值保留
    b.quit()
    assert app.DisplayAlerts is True


def test_excel_alerts_scope_restores_on_exception(monkeypatch):
    app = types.SimpleNamespace(DisplayAlerts=True)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_app_returning(app))
    b = excel.ExcelApp()
    with pytest.raises(RuntimeError), b._alerts_scope():
        assert app.DisplayAlerts is False
        raise RuntimeError("boom")
    assert app.DisplayAlerts is True


# --- P1-6：save_pdf 走 ExportAsFixedFormat（官方导出 API） ---


class _FakePresExport:
    def __init__(self):
        self.calls = []

    def ExportAsFixedFormat(self, path, *a, **k):
        self.calls.append((path, a, k))


class _FakeDocExport:
    def __init__(self):
        self.calls = []

    def ExportAsFixedFormat(self, path, *a, **k):
        self.calls.append((path, a, k))


def test_ppt_save_pdf_uses_export_as_fixed_format(tmp_path):
    target = tmp_path / "x.pdf"
    p = ppt.PptApp.__new__(ppt.PptApp)
    p.app = types.SimpleNamespace(DisplayAlerts=0)
    fake = _FakePresExport()
    p.active_pres = lambda doc_id=None: fake
    p.save_pdf(str(target))
    path, args, kwargs = fake.calls[0]
    assert path == os.path.abspath(str(target))
    assert kwargs["Intent"] == 2  # ppFixedFormatIntentPrint
    assert kwargs["OutputType"] == ppt.PP_FIXED_FORMAT_TYPE_PDF


def test_word_save_pdf_uses_export_as_fixed_format(tmp_path):
    target = tmp_path / "x.pdf"
    d = word.WordApp.__new__(word.WordApp)
    d.app = types.SimpleNamespace(DisplayAlerts=0)
    fake = _FakeDocExport()
    d.active_doc = lambda doc_id=None: fake
    d.save_pdf(str(target))
    path, args, kwargs = fake.calls[0]
    assert path == os.path.abspath(str(target))
    assert kwargs["ExportFormat"] == word.WD_EXPORT_FORMAT_PDF
