"""office check：环境就绪诊断（只读，不拉起 Office/不启动 server）。

每个检查返回 Check；main() 汇总渲染文本或 --json。
顶层不 import pywin32/playwright/core：全部函数内惰性 import + 平台守卫，
保证无 Office / 非 Windows / 无 chromium 也能跑完并优雅降级。
"""

from __future__ import annotations

import importlib
import importlib.metadata
import json
import platform
import shutil
import sys
from dataclasses import dataclass

from . import __version__
from .exceptions import UnsupportedPlatformError

# (dist 名, import 名, 适用 sys.platform 或 None)：pywin32 仅 Windows，
# 非 Windows 上 win32com import 不存在（Linux 纯模块 CI 不能炸）。
# 其余依赖跨平台适用。
_DEPS = [
    ("pywin32", "win32com", "win32"),
    ("python-pptx", "pptx", None),
    ("lxml", "lxml", None),
    ("fonttools", "fontTools", None),  # dist 名小写，import 名大写 T
    ("playwright", "playwright", None),
    ("Pillow", "PIL", None),
    ("mcp", "mcp", None),
]


def _platform_deps() -> list[tuple[str, str]]:
    """当前平台实际适用的依赖清单（保持声明顺序）。"""
    return [
        (dist, mod) for dist, mod, sys_name in _DEPS if sys_name is None or sys.platform == sys_name
    ]


_OFFICE = [
    ("word", "Word.Application"),
    ("excel", "Excel.Application"),
    ("ppt", "PowerPoint.Application"),
]
_PY_REQ = (3, 10)


@dataclass
class Check:
    section: str
    name: str
    ok: bool
    detail: str
    hint: str = ""
    warn: bool = False  # True：只计警告，不影响 exit code


def _check_python() -> Check:
    cur = sys.version_info
    ok = (cur.major, cur.minor) >= _PY_REQ
    return Check(
        "运行时",
        "Python",
        ok,
        f"{cur.major}.{cur.minor}.{cur.micro}（要求 ≥ {_PY_REQ[0]}.{_PY_REQ[1]}）",
        hint=("请安装 Python 3.10+" if not ok else ""),
    )


def _check_platform() -> Check:
    if platform.system() == "Windows":
        return Check("运行时", "平台", True, "Windows")
    return Check("运行时", "平台", False, "非 Windows，Office/COM 检查跳过", warn=True)


def _check_offipy() -> Check:
    return Check("offipy", "版本", True, __version__)


# dist 名 → 所属 extra 的安装提示（offipy check 报告缺失 extra，P1-5）
_EXTRA_HINT = {
    "pywin32": "uv pip install 'offipy[office]'",
    "python-pptx": "uv pip install 'offipy[deck]'",
    "lxml": "uv pip install 'offipy[deck]'",
    "fonttools": "uv pip install 'offipy[deck]'",
    "playwright": "uv pip install 'offipy[deck]'",
    "Pillow": "uv pip install 'offipy[deck]'",
    "mcp": "uv pip install 'offipy[mcp]'",
}


def _check_dependencies() -> list[Check]:
    out = []
    for dist, mod in _platform_deps():
        try:
            importlib.import_module(mod)
            version = importlib.metadata.version(dist)
            out.append(Check("依赖", dist, True, version))
        except ImportError:
            out.append(
                Check(
                    "依赖",
                    dist,
                    False,
                    "未安装",
                    hint=_EXTRA_HINT.get(dist, f"uv pip install {dist}"),
                )
            )
    return out


def _office_installed(progid: str) -> bool:
    """注册表查 ProgID（HKLM/HKCU × 64/32 位视图），无副作用、不拉起应用。"""
    import winreg

    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for base in (r"SOFTWARE\Classes", r"SOFTWARE\WOW6432Node\Classes"):
            try:
                with winreg.OpenKey(root, base + "\\" + progid):
                    return True
            except OSError:
                continue
    return False


def _check_office() -> list[Check]:
    if platform.system() != "Windows":
        return [Check("Office 套件", "Word/Excel/PowerPoint", False, "仅 Windows 支持", warn=True)]
    try:
        from . import core  # 惰性 import；core 顶层不 import pywin32

        running = {app: core.running(app) for app, _ in _OFFICE}
    except (ImportError, UnsupportedPlatformError):
        running = {}
    out = []
    for app, progid in _OFFICE:
        installed = _office_installed(progid)
        if installed:
            state = "已安装，当前有存活实例" if running.get(app) else "已安装"
            out.append(Check("Office 套件", app, True, state))
        else:
            out.append(
                Check("Office 套件", app, False, "未检测到安装", hint="请安装 Microsoft Office")
            )
    return out


def _check_browser() -> Check:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return Check(
            "浏览器", "Chromium", False, "playwright 包未安装", hint="uv pip install 'offipy[deck]'"
        )
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
        return Check("浏览器", "Chromium", True, "可正常启动（headless）")
    except Exception as e:
        # playwright.sync_api.Error（chromium 缺失/无法启动），宽捕获做降级
        return Check(
            "浏览器", "Chromium", False, f"启动失败: {e}", hint="uv run playwright install chromium"
        )


def _check_server() -> Check:
    from .client import _probe

    state = _probe()
    if state == "ok":
        return Check("本地 server", "127.0.0.1:8890", True, "运行中")
    if state == "auth_fail":
        return Check(
            "本地 server", "127.0.0.1:8890", False, "端口有 server 但 token 不匹配", warn=True
        )
    if state == "mismatch":
        return Check("本地 server", "127.0.0.1:8890", False, "端口被非 offipy 进程占用", warn=True)
    return Check(
        "本地 server",
        "127.0.0.1:8890",
        False,
        "未运行（首次调用 office <app> <op> 会自动拉起）",
        warn=True,
    )


def _check_pdf() -> Check:
    soffice = shutil.which("soffice")
    try:
        import pdf2image  # noqa: F401

        pdf = "pdf2image 可用"
    except ImportError:
        pdf = "pdf2image 未安装"
    detail = f"LibreOffice: {'有' if soffice else '无'}；{pdf}"
    return Check("PDF 可选路径", "LibreOffice/pdf2image", True, detail, warn=True)


# check --profile 过滤表：office/deck/mcp 都以 core 为基线，再叠加各自分组
_CORE_SECTIONS = {"运行时", "offipy", "本地 server", "PDF 可选路径"}
_PROFILE_SECTIONS = {
    "office": {"Office 套件"},
    "deck": {"浏览器"},
    "mcp": set(),
}
_PROFILE_DEPS = {
    "office": {"pywin32"},
    "deck": {"python-pptx", "lxml", "fonttools", "playwright", "Pillow"},
    "mcp": {"mcp"},
}


def run(profile: str | None = None) -> list[Check]:
    checks = [_check_python(), _check_platform(), _check_offipy()]
    if profile is None:
        checks += _check_dependencies()
        checks += _check_office()
        checks.append(_check_browser())
        checks.append(_check_server())
        checks.append(_check_pdf())
        return checks
    # profile 模式按需构建，跳过无关的昂贵检查（如非 deck 不启 chromium）
    sections = _CORE_SECTIONS | _PROFILE_SECTIONS.get(profile, set())
    deps = _PROFILE_DEPS.get(profile, set())
    if deps:
        checks += [c for c in _check_dependencies() if c.name in deps]
    if "Office 套件" in sections:
        checks += _check_office()
    if "浏览器" in sections:
        checks.append(_check_browser())
    checks.append(_check_server())
    checks.append(_check_pdf())
    return checks


def _mark(c: Check) -> str:
    if c.warn:
        return "⚠"
    return "✓" if c.ok else "✗"


def render_text(checks: list[Check]) -> str:
    lines = [f"[offipy 环境检查] v{__version__}"]
    section = None
    for c in checks:
        if c.section != section:
            lines.append("")
            lines.append(c.section)
            section = c.section
        lines.append(f"  {_mark(c)} {c.name}: {c.detail}")
        if not c.ok and c.hint:
            lines.append(f"      修复: {c.hint}")
    fails = sum(1 for c in checks if not c.ok and not c.warn)
    warns = sum(1 for c in checks if c.warn)
    passed = len(checks) - fails - warns
    lines.append("")
    lines.append(f"通过 {passed} | 警告 {warns} | 失败 {fails}")
    lines.append("结果: " + ("环境就绪 ✓" if fails == 0 else f"存在 {fails} 项异常（见上方 ✗）"))
    return "\n".join(lines)


def render_json(checks: list[Check]) -> str:
    fails = sum(1 for c in checks if not c.ok and not c.warn)
    warns = sum(1 for c in checks if c.warn)
    payload = {
        "version": __version__,
        "ok": fails == 0,
        "fails": fails,
        "warns": warns,
        "checks": [
            {
                "section": c.section,
                "name": c.name,
                "ok": c.ok,
                "warn": c.warn,
                "detail": c.detail,
                "hint": c.hint,
            }
            for c in checks
        ],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def main(json_output: bool = False, profile: str | None = None) -> int:
    checks = run(profile)
    if json_output:
        print(render_json(checks))
    else:
        print(render_text(checks))
    return 0 if sum(1 for c in checks if not c.ok and not c.warn) == 0 else 1
