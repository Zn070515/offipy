"""pytest 共享配置。

COM 集成测试（操作真实 Office）要求本线程先 CoInitialize。这里在 conftest
导入时初始化一次——与 server 的「单线程初始化一次」哲学一致，也让测试模块
顶层的 `skipif(not core.running(...))` 能安全执行。

pywin32 仅支持 Windows：缺失时 pythoncom 置 None，live_ppt 等 COM 依赖的
fixture 直接跳过，保证非 Windows / 未装 pywin32 也能跑完整套 collection。
"""

import pytest

try:
    import pythoncom
except ImportError:  # pragma: no cover - 非 Windows / 未装 pywin32
    pythoncom = None

from offipy import core

if pythoncom is not None:
    pythoncom.CoInitialize()


@pytest.fixture
def live_ppt():
    """需要存活 PowerPoint 时使用；缺实例则跳过（session server 8890 持有）。"""
    if pythoncom is None or not core.running("ppt"):
        pytest.skip("需要存活 PowerPoint（session server 8890 持有）")
    return True
