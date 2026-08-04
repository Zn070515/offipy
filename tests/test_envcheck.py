"""envcheck（office check）纯逻辑测试：mock 注入，不依赖 Office/COM/浏览器。

envcheck 顶层不 import pywin32/playwright/core，全部检查函数惰性 import，
因此这些测试用 monkeypatch 注入假 winreg / playwright / core.running / _ping。
"""

import builtins
import json
import sys
import types
from collections import namedtuple

from offipy import envcheck

# --- 运行时 ---


def test_check_python_ok():
    c = envcheck._check_python()
    assert c.section == "运行时"
    assert c.ok == (sys.version_info >= (3, 10))


def test_check_python_old_fails(monkeypatch):
    Ver = namedtuple("Ver", "major minor micro")
    monkeypatch.setattr(envcheck.sys, "version_info", Ver(3, 9, 5))
    c = envcheck._check_python()
    assert c.ok is False
    assert "3.9.5" in c.detail
    assert "安装 Python 3.10+" in c.hint


def test_check_platform_windows_ok():
    c = envcheck._check_platform()
    assert c.section == "运行时"
    assert c.ok == (sys.platform == "win32")


def test_check_platform_non_windows_warns(monkeypatch):
    monkeypatch.setattr(envcheck.platform, "system", lambda: "Linux")
    c = envcheck._check_platform()
    assert c.ok is False
    assert c.warn is True
    assert "非 Windows" in c.detail


def test_check_offipy_version():
    c = envcheck._check_offipy()
    assert c.ok is True
    assert c.detail == envcheck.__version__


# --- 依赖 ---


def test_check_dependencies_all_ok():
    checks = envcheck._check_dependencies()
    assert len(checks) == len(envcheck._DEPS)
    for c in checks:
        assert c.ok is True
        assert c.detail  # 版本号非空


def test_check_dependencies_import_failure(monkeypatch):
    import importlib as _importlib

    real_import_module = _importlib.import_module

    def fake_import_module(name, *args, **kwargs):
        if name == "fontTools":
            raise ImportError("not installed")
        return real_import_module(name, *args, **kwargs)

    monkeypatch.setattr(_importlib, "import_module", fake_import_module)
    checks = {c.name: c for c in envcheck._check_dependencies()}
    assert checks["fonttools"].ok is False
    assert "uv pip install fonttools" in checks["fonttools"].hint
    assert checks["pywin32"].ok is True


# --- Office 安装检测（fake winreg） ---


class _Ctx:
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class _FakeWinreg:
    HKEY_LOCAL_MACHINE = object()
    HKEY_CURRENT_USER = object()
    existing = {
        r"SOFTWARE\Classes\Word.Application",
        r"SOFTWARE\Classes\Excel.Application",
    }

    @staticmethod
    def OpenKey(root, path):
        if path in _FakeWinreg.existing:
            return _Ctx()
        raise OSError("path not found")


def test_office_installed_hit(monkeypatch):
    monkeypatch.setitem(sys.modules, "winreg", _FakeWinreg)
    assert envcheck._office_installed("Word.Application") is True
    assert envcheck._office_installed("Excel.Application") is True


def test_office_installed_miss(monkeypatch):
    monkeypatch.setitem(sys.modules, "winreg", _FakeWinreg)
    assert envcheck._office_installed("PowerPoint.Application") is False


# --- Office 套件检查 ---


def test_check_office_installed_with_running(monkeypatch):
    monkeypatch.setattr(envcheck.platform, "system", lambda: "Windows")
    monkeypatch.setattr(envcheck, "_office_installed", lambda progid: True)
    running = {"word": True, "excel": False, "ppt": True}
    monkeypatch.setattr("offipy.core.running", lambda app: running.get(app))
    checks = envcheck._check_office()
    assert len(checks) == 3
    assert checks[0].detail == "已安装，当前有存活实例"  # word 存活
    assert checks[1].detail == "已安装"  # excel 无存活实例


def test_check_office_not_installed(monkeypatch):
    monkeypatch.setattr(envcheck.platform, "system", lambda: "Windows")
    monkeypatch.setattr(envcheck, "_office_installed", lambda progid: False)
    checks = envcheck._check_office()
    assert checks[0].ok is False
    assert checks[0].hint == "请安装 Microsoft Office"


def test_check_office_non_windows(monkeypatch):
    monkeypatch.setattr(envcheck.platform, "system", lambda: "Linux")
    checks = envcheck._check_office()
    assert len(checks) == 1
    assert checks[0].ok is False
    assert checks[0].warn is True


# --- 浏览器检查（fake playwright.sync_api） ---


def _inject_fake_playwright(monkeypatch, playwright_ctx):
    mod = types.ModuleType("playwright.sync_api")
    mod.sync_playwright = lambda: playwright_ctx
    monkeypatch.setitem(sys.modules, "playwright.sync_api", mod)


class _OkChromium:
    def launch(self, headless=False):
        assert headless is True
        return self

    def close(self):
        return None


class _OkPW:
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    @property
    def chromium(self):
        return _OkChromium()


class _FailPW:
    def __enter__(self):
        raise RuntimeError("Executable doesn't exist at ...")

    def __exit__(self, *exc):
        return False


def test_check_browser_ok(monkeypatch):
    _inject_fake_playwright(monkeypatch, _OkPW())
    c = envcheck._check_browser()
    assert c.ok is True
    assert "headless" in c.detail


def test_check_browser_launch_fails(monkeypatch):
    _inject_fake_playwright(monkeypatch, _FailPW())
    c = envcheck._check_browser()
    assert c.ok is False
    assert "启动失败" in c.detail
    assert "playwright install chromium" in c.hint


def test_check_browser_missing_package(monkeypatch):
    # 让 from playwright.sync_api import sync_playwright 失败（playwright 未装）
    monkeypatch.delitem(sys.modules, "playwright.sync_api", raising=False)
    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "playwright.sync_api":
            raise ImportError("no playwright")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    c = envcheck._check_browser()
    assert c.ok is False
    assert c.hint == "uv pip install playwright"


# --- 本地 server ---


def test_check_server_running(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")
    c = envcheck._check_server()
    assert c.ok is True
    assert c.detail == "运行中"


def test_check_server_not_running_warns(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "down")
    c = envcheck._check_server()
    assert c.ok is False
    assert c.warn is True


def test_check_server_auth_fail_warns(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "auth_fail")
    c = envcheck._check_server()
    assert c.ok is False
    assert "token 不匹配" in c.detail


# --- PDF 可选路径 ---


def test_check_pdf_both_present(monkeypatch):
    monkeypatch.setattr(envcheck.shutil, "which", lambda name: "C:/soft/soffice.exe")
    monkeypatch.setitem(sys.modules, "pdf2image", types.ModuleType("pdf2image"))
    c = envcheck._check_pdf()
    assert c.warn is True
    assert "LibreOffice: 有" in c.detail
    assert "pdf2image 可用" in c.detail


def test_check_pdf_both_missing(monkeypatch):
    monkeypatch.setattr(envcheck.shutil, "which", lambda name: None)
    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "pdf2image":
            raise ImportError("no pdf2image")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    c = envcheck._check_pdf()
    assert c.warn is True
    assert "LibreOffice: 无" in c.detail
    assert "pdf2image 未安装" in c.detail


# --- 渲染与 exit code ---


def _checks_fixture():
    return [
        envcheck.Check("运行时", "Python", True, "3.12"),
        envcheck.Check("依赖", "pywin32", False, "未安装", hint="uv pip install pywin32"),
        envcheck.Check("本地 server", "127.0.0.1:8890", False, "未运行", warn=True),
    ]


def test_render_text_marks():
    text = envcheck.render_text(_checks_fixture())
    assert "✓ Python" in text
    assert "✗ pywin32" in text
    assert "⚠ 127.0.0.1:8890" in text
    assert "修复: uv pip install pywin32" in text
    assert "通过 1 | 警告 1 | 失败 1" in text


def test_render_text_section_grouping():
    text = envcheck.render_text(_checks_fixture())
    # 分组标题按 section 出现，且带版本头部
    assert "[offipy 环境检查]" in text
    assert "\n运行时\n" in text
    assert "\n依赖\n" in text
    assert "\n本地 server\n" in text


def test_render_json_structure():
    data = json.loads(envcheck.render_json(_checks_fixture()))
    assert data["ok"] is False
    assert data["fails"] == 1
    assert data["warns"] == 1
    assert data["version"] == envcheck.__version__
    assert len(data["checks"]) == 3
    assert data["checks"][0]["name"] == "Python"
    assert data["checks"][1]["hint"] == "uv pip install pywin32"


def test_main_exit_one_on_hard_fail(monkeypatch, capsys):
    monkeypatch.setattr(envcheck, "run", lambda: _checks_fixture())
    assert envcheck.main() == 1
    assert "结果:" in capsys.readouterr().out


def test_main_exit_zero_when_only_warns(monkeypatch, capsys):
    only_warns = [
        envcheck.Check("运行时", "Python", True, "3.12"),
        envcheck.Check("本地 server", "x", False, "未运行", warn=True),
    ]
    monkeypatch.setattr(envcheck, "run", lambda: only_warns)
    assert envcheck.main() == 0


def test_main_json_output(monkeypatch, capsys):
    all_ok = [envcheck.Check("运行时", "Python", True, "3.12")]
    monkeypatch.setattr(envcheck, "run", lambda: all_ok)
    assert envcheck.main(json_output=True) == 0
    data = json.loads(capsys.readouterr().out)
    assert data["ok"] is True
