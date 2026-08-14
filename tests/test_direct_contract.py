"""P1-2/P1-4 测试：ensure_app 可见性契约 + direct 线程/COM 套间。

ensure_app 连到既有实例默认不改其可见性（用户正用的窗口不被抢改）；
modify_existing_visibility=True 才 _set_visible。com_apartment 提供 STA
套间上下文，pythoncom 不可用时是 no-op。
"""

import sys
import types

import pytest

from offipy import core
from offipy.direct import com_apartment

# --- P1-2：ensure_app 可见性契约 ---


def test_ensure_app_attached_preserves_visibility(monkeypatch):
    app = types.SimpleNamespace(Visible=True)
    monkeypatch.setattr("offipy.core.connect", lambda app_name: app)
    calls = []
    monkeypatch.setattr(
        "offipy.core._set_visible", lambda obj, visible: calls.append(f"visible:{visible}")
    )
    monkeypatch.setattr("offipy.core._set_usercontrol", lambda obj: calls.append("usercontrol"))
    obj, created = core.ensure_app("excel", visible=False)
    assert created is False
    assert obj is app
    # 不改既有实例可见性，但 usercontrol 保留（实例随 Python 存活）
    assert calls == ["usercontrol"]


def test_ensure_app_attached_modify_existing_visibility(monkeypatch):
    app = types.SimpleNamespace(Visible=True)
    monkeypatch.setattr("offipy.core.connect", lambda app_name: app)
    calls = []
    monkeypatch.setattr(
        "offipy.core._set_visible", lambda obj, visible: calls.append(f"visible:{visible}")
    )
    monkeypatch.setattr("offipy.core._set_usercontrol", lambda obj: calls.append("usercontrol"))
    _obj, created = core.ensure_app("excel", visible=False, modify_existing_visibility=True)
    assert created is False
    assert calls == ["visible:False", "usercontrol"]  # 显式要求才改可见性


def test_ensure_app_launch_sets_visible(monkeypatch):
    app = types.SimpleNamespace(Visible=True)
    calls = []
    monkeypatch.setattr("offipy.core.connect", lambda app_name: None)  # 无既有实例
    monkeypatch.setattr(
        "offipy.core.launch", lambda app_name, visible: calls.append(f"visible:{visible}") or app
    )
    monkeypatch.setattr("offipy.core._set_usercontrol", lambda obj: None)
    _obj, created = core.ensure_app("word", visible=True)
    assert created is True
    assert calls == ["visible:True"]


# --- P1-4：com_apartment STA 套间 ---


def test_com_apartment_noop_without_pythoncom(monkeypatch):
    def _raise_import(name, *a, **k):
        if name == "pythoncom":
            raise ImportError("no pythoncom")
        return __import__(name, *a, **k)

    monkeypatch.setattr("builtins.__import__", _raise_import)
    with com_apartment():
        pass  # 非 Windows / 无 pythoncom → no-op，enter/exit 无异常


def test_com_apartment_initializes_and_uninitializes(monkeypatch):
    calls = []
    fake = types.ModuleType("pythoncom")
    fake.CoInitialize = lambda: calls.append("init")
    fake.CoUninitialize = lambda: calls.append("uninit")
    monkeypatch.setitem(sys.modules, "pythoncom", fake)
    with com_apartment():
        assert calls == ["init"]
    assert calls == ["init", "uninit"]


def test_com_apartment_uninitializes_on_exception(monkeypatch):
    calls = []
    fake = types.ModuleType("pythoncom")
    fake.CoInitialize = lambda: calls.append("init")
    fake.CoUninitialize = lambda: calls.append("uninit")
    monkeypatch.setitem(sys.modules, "pythoncom", fake)
    with pytest.raises(RuntimeError), com_apartment():
        raise RuntimeError("boom")
    assert calls == ["init", "uninit"]  # 异常路径也成对还原
