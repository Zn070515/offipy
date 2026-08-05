"""跨平台惰性 COM 测试：import offipy 不拉 pywin32，非 Windows 抛语义化异常。"""

import subprocess
import sys

import pytest

from offipy import core
from offipy.exceptions import ComOperationError, OfficeUnavailableError, UnsupportedPlatformError


def test_import_offipy_in_clean_interpreter_pulls_no_pywin32():
    """P0-7 回归：干净解释器里 import offipy 不触发 win32com/pythoncom。"""
    code = (
        "import sys\n"
        "import offipy\n"
        "assert 'win32com' not in sys.modules, 'win32com 被顶层 import'\n"
        "assert 'pythoncom' not in sys.modules, 'pythoncom 被顶层 import'\n"
        "print(offipy.__version__)\n"
    )
    out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=60)
    assert out.returncode == 0, out.stderr


def test_com_raises_unsupported_on_non_windows(monkeypatch):
    # 清掉模块级缓存，否则 Windows 上已加载的 _COM 会让 _com() 早退、跳过平台检查
    monkeypatch.setattr(core, "_COM", None)
    monkeypatch.setattr(sys, "platform", "linux")
    with pytest.raises(UnsupportedPlatformError):
        core._com()


def test_connect_uses_lazy_com_bundle(monkeypatch):
    class FakePywintypes:
        com_error = ValueError

    class FakeWin32:
        def GetActiveObject(self, progid):
            return "obj"

    bundle = core._ComBundle(pywintypes=FakePywintypes(), win32com=FakeWin32(), gencache=None)
    monkeypatch.setattr(core, "_com", lambda: bundle)
    assert core.connect("ppt") == "obj"


def test_connect_none_when_com_missing(monkeypatch):
    # P1-4 COM HRESULT 分类：仅「未运行/类未注册」两类 HRESULT → None（触发 launch）
    class FakeComError(Exception):
        hresult = -2147221021  # MK_E_UNAVAILABLE：ROT 无存活对象

    class FakePywintypes:
        com_error = FakeComError

    class FakeWin32:
        def GetActiveObject(self, progid):
            raise FakePywintypes.com_error()

    bundle = core._ComBundle(pywintypes=FakePywintypes(), win32com=FakeWin32(), gencache=None)
    monkeypatch.setattr(core, "_com", lambda: bundle)
    assert core.connect("ppt") is None


def test_connect_none_when_class_string_invalid(monkeypatch):
    # CI windows runner 无 Office：GetActiveObject 抛 CO_E_CLASSSTRING（Invalid class string）
    class FakeComError(Exception):
        hresult = -2147221005  # CO_E_CLASSSTRING：ProgID 无法映射到 CLSID（Office 未安装）

    class FakePywintypes:
        com_error = FakeComError

    class FakeWin32:
        def GetActiveObject(self, progid):
            raise FakePywintypes.com_error()

    bundle = core._ComBundle(pywintypes=FakePywintypes(), win32com=FakeWin32(), gencache=None)
    monkeypatch.setattr(core, "_com", lambda: bundle)
    assert core.connect("ppt") is None


def test_connect_raises_on_permission_com_error(monkeypatch):
    # P1-4：非「未运行」HRESULT（如权限拒绝）→ 抛 ComOperationError，绝不静默拉起
    class FakeComError(Exception):
        hresult = 0x80070005  # E_ACCESSDENIED

    class FakePywintypes:
        com_error = FakeComError

    class FakeWin32:
        def GetActiveObject(self, progid):
            raise FakePywintypes.com_error()

    bundle = core._ComBundle(pywintypes=FakePywintypes(), win32com=FakeWin32(), gencache=None)
    monkeypatch.setattr(core, "_com", lambda: bundle)
    with pytest.raises(ComOperationError):
        core.connect("ppt")


def test_ensure_app_raises_office_unavailable_when_launch_fails(monkeypatch):
    monkeypatch.setattr(core, "connect", lambda app: None)

    def boom(app, visible=True):
        raise RuntimeError("COM 初始化失败")

    monkeypatch.setattr(core, "launch", boom)
    with pytest.raises(OfficeUnavailableError):
        core.ensure_app("ppt")
