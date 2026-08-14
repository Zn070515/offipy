"""提交 0：三个 App 的 activate 未知句柄统一会话边界提示。

0.10.2 已统一 open 类路由 6 处消息；本测试收口 activate(doc_id) 的未知句柄
错误语义——doc_id 只在同会话内有效，本地直连与会话式 Remote*/CLI/HTTP 互不相通。

用 `__new__` 绕过 __init__，注入 fake 文档表与存活探针，纯逻辑验证。
"""

import pytest

from offipy import core, excel, ppt, word
from offipy.exceptions import TargetNotFoundError

APPS = [
    pytest.param(excel, "book", id="excel"),
    pytest.param(word, "doc", id="word"),
    pytest.param(ppt, "pres", id="ppt"),
]


class _LiveDoc:
    """带 Activate() 的存活假体；Windows 故意缺失以走 Ppt 的回退分支。"""

    def __init__(self, name):
        self.Name = name
        self.activated = 0

    def Activate(self):
        self.activated += 1

    @property
    def Windows(self):
        raise AttributeError("no windows (Ppt falls back to Pres.Activate)")


_APP_CLS = {excel: excel.ExcelApp, word: word.WordApp, ppt: ppt.PptApp}


def _make_app(monkeypatch, module):
    monkeypatch.setattr(core, "doc_alive", lambda *a: True)
    app = _APP_CLS[module].__new__(_APP_CLS[module])
    app.app = object()
    app._docs = {}
    app._active_id = None
    return app


@pytest.mark.parametrize(("module", "prefix"), APPS)
def test_activate_unknown_id_message_session_boundary(monkeypatch, module, prefix):
    app = _make_app(monkeypatch, module)
    with pytest.raises(TargetNotFoundError) as ei:
        app.activate("nope")
    msg = str(ei.value)
    assert "同会话" in msg
    assert "list_docs" in msg


@pytest.mark.parametrize(("module", "prefix"), APPS)
def test_activate_valid_id_works(monkeypatch, module, prefix):
    app = _make_app(monkeypatch, module)
    live = _LiveDoc(f"{prefix}1")
    app._docs[f"{prefix}1"] = live
    assert app.activate(f"{prefix}1") == f"{prefix}1"
    assert live.activated == 1
    assert app._active_id == f"{prefix}1"
