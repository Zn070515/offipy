"""C1（H5）: quit()/quit_app 在 Quit() 前必须抑制弹窗，防脏文档模态框挂死。

各 app 的 DisplayAlerts「抑制全部弹窗」值不同：ppt=PP_ALERTS_NONE(1)、
word=WD_ALERTS_NONE(0)、excel=False。Quit 前还原成 _saved_alerts 会让
「是否保存」模态对话框弹出 → COM 调用阻塞 → 进程挂死。
"""

import pytest

from offipy import core, excel, ppt, word

# 测试的是 Windows Office COM 的 quit/DisplayAlerts 语义：构造 PptApp/WordApp/ExcelApp
# 会走 core._com()（仅 Windows）。ubuntu 纯模块子集（-m "not com"）必须排除；
# Windows 全量下这些测试纯 mock 可跑，不需要存活 Office。
pytestmark = [pytest.mark.com]


class _AlertRecorder:
    """记录 DisplayAlerts 赋值序列与 Quit 执行时的值。"""

    def __init__(self, saved):
        self.saved = saved
        self._value = saved
        self.assignments = []
        self.quit_value = None
        self.fail = False

    @property
    def DisplayAlerts(self):
        return self._value

    @DisplayAlerts.setter
    def DisplayAlerts(self, v):
        self._value = v
        self.assignments.append(v)

    def Quit(self):
        self.quit_value = self._value
        if self.fail:
            raise RuntimeError("RPC_E_DISCONNECTED")


def _ensure_owned(app):
    def _ensure(app_name, visible=True, modify_existing_visibility=False):
        return app, True

    return _ensure


def _ensure_attached(app):
    def _ensure(app_name, visible=True, modify_existing_visibility=False):
        return app, False

    return _ensure


def _quit_setup(monkeypatch, recorder):
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_owned(recorder))
    monkeypatch.setattr("offipy.core.quit_app", lambda n: True)
    monkeypatch.setattr("offipy.core.app_process_pid", lambda obj, app: 123)
    monkeypatch.setattr("offipy.core.wait_process_exit", lambda pid, timeout=2.0: True)
    monkeypatch.setattr("offipy.core.reap_process", lambda pid: None)


@pytest.mark.parametrize(
    ("cls", "none_val", "saved"),
    [
        (ppt.PptApp, ppt.PP_ALERTS_NONE, 0),
        (word.WordApp, word.WD_ALERTS_NONE, -1),
        (excel.ExcelApp, False, True),
    ],
)
def test_quit_suppresses_alerts_before_quit(monkeypatch, cls, none_val, saved):
    rec = _AlertRecorder(saved)
    _quit_setup(monkeypatch, rec)
    obj = cls()
    assert obj.quit() is None
    # Quit() 执行瞬间 DisplayAlerts 必须已是抑制值（否则脏文档弹保存框挂死）
    assert rec.quit_value == none_val
    # Quit 之后 finally 还原 _saved_alerts（失败路径仍还给用户原值）
    assert rec.assignments == [none_val, saved]
    assert rec.DisplayAlerts == saved


@pytest.mark.parametrize(
    ("cls", "none_val", "saved"),
    [
        (ppt.PptApp, ppt.PP_ALERTS_NONE, 0),
        (word.WordApp, word.WD_ALERTS_NONE, 1),
        (excel.ExcelApp, False, False),
    ],
)
def test_quit_exception_process_exited_restores_alerts(monkeypatch, cls, none_val, saved):
    rec = _AlertRecorder(saved)
    rec.fail = True  # Quit 抛断连错误
    _quit_setup(monkeypatch, rec)
    monkeypatch.setattr("offipy.core.doc_alive", lambda obj: False)  # 进程已退出
    obj = cls()
    assert obj.quit() is True
    assert rec.quit_value == none_val  # 抛错瞬间已处于抑制态
    assert rec.DisplayAlerts == saved  # finally 还原


@pytest.mark.parametrize(
    ("cls", "none_val"),
    [
        (ppt.PptApp, ppt.PP_ALERTS_NONE),
        (word.WordApp, word.WD_ALERTS_NONE),
        (excel.ExcelApp, False),
    ],
)
def test_quit_attached_force_also_suppresses(monkeypatch, cls, none_val):
    # force=True 连既有实例也直接退：同样必须抑制弹窗
    rec = _AlertRecorder(0)
    monkeypatch.setattr("offipy.core.ensure_app", _ensure_attached(rec))
    monkeypatch.setattr("offipy.core.quit_app", lambda n: True)
    monkeypatch.setattr("offipy.core.app_process_pid", lambda obj, app: 123)
    monkeypatch.setattr("offipy.core.wait_process_exit", lambda pid, timeout=2.0: True)
    obj = cls()
    assert obj.quit(force=True) is None
    assert rec.quit_value == none_val


def test_quit_app_suppresses_alerts(monkeypatch):
    # core.quit_app 同样在 Quit 前抑制弹窗（Excel 用 False）
    rec = _AlertRecorder(True)
    monkeypatch.setattr("offipy.core.connect", lambda app: rec)
    monkeypatch.setattr("offipy.core.app_process_pid", lambda obj, app: 55)
    monkeypatch.setattr("offipy.core.wait_process_exit", lambda pid, timeout=2.0: True)
    assert core.quit_app("excel") is True
    assert rec.quit_value is False
