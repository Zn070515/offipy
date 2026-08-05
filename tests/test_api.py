"""高层 API facade 与会话语义（P1.2）测试：全 mock，不触 COM。

覆盖：
- Excel()/Word()/Ppt() 上下文管理器进出
- 未显式定义的 op 经 __getattr__ 代理到底层 app
- offipy 异常透传
- active_doc 实时优先 / 缓存回退 / 死缓存穿透 / 无文档不隐式创建（P0-8）
"""

import pytest

from offipy import api, exceptions


def test_facade_context_manager_proxies_op(monkeypatch):
    captured = {}

    class FakeApp:
        def new_book(self):
            captured["called"] = True
            return "book"

    monkeypatch.setattr("offipy.api.ExcelApp", lambda visible=True: FakeApp())
    with api.Excel() as x:
        assert x.new_book() == "book"
    assert captured["called"]


def test_facade_quit_proxies(monkeypatch):
    captured = {}

    class FakeApp:
        def quit(self):
            captured["quit"] = True

    monkeypatch.setattr("offipy.api.WordApp", lambda visible=True: FakeApp())
    api.Word().quit()
    assert captured["quit"]


def test_facade_enter_exit_no_com(monkeypatch):
    monkeypatch.setattr("offipy.api.PptApp", lambda visible=True: object())
    with api.Ppt() as p:
        assert p._app is not None


def test_facade_propagates_offipy_exception(monkeypatch):
    class Boom:
        def add_slide(self):
            raise exceptions.ConversionError("boom")

    monkeypatch.setattr("offipy.api.PptApp", lambda visible=True: Boom())
    with pytest.raises(exceptions.ConversionError, match="boom"), api.Ppt() as p:
        p.add_slide()


def test_facade_missing_op_raises_attribute_error(monkeypatch):
    class FakeApp:
        pass

    monkeypatch.setattr("offipy.api.ExcelApp", lambda visible=True: FakeApp())
    with pytest.raises(AttributeError):
        api.Excel().no_such_op()


def test_active_doc_prefers_live(monkeypatch):
    from offipy import core
    from offipy.excel import ExcelApp

    live = object()
    app = ExcelApp.__new__(ExcelApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    monkeypatch.setattr(core, "active_doc", lambda name, attr: live)
    assert app.active_book() is live
    assert live in app._docs.values()  # 实时句柄并入文档表并设为活动


def test_active_doc_stale_registry_not_used_without_real_active(monkeypatch):
    # P0-3 doc_id 权威：缺省解析一律实时读真实 Active，绝不静默回退到陈旧的
    # _active_id 登记——实时解析不到就是 None（防「用户看到 B、Agent 以为 A」）。
    from offipy import core
    from offipy.excel import ExcelApp

    cached = object()
    fake_app = type("F", (), {"ActiveWorkbook": None})()
    app = ExcelApp.__new__(ExcelApp)
    app._docs = {"book1": cached}
    app._active_id = "book1"
    app.app = fake_app
    monkeypatch.setattr(core, "active_doc", lambda name, attr: None)
    monkeypatch.setattr(core, "doc_alive", lambda obj: True)
    assert app.active_book() is None


def test_active_doc_dead_cache_falls_through(monkeypatch):
    from offipy import core
    from offipy.excel import ExcelApp

    live_book = object()
    fake_app = type("F", (), {"ActiveWorkbook": live_book})()
    app = ExcelApp.__new__(ExcelApp)
    app._docs = {"book1": object()}  # 死缓存
    app._active_id = "book1"
    app._seq = 0
    app.app = fake_app
    monkeypatch.setattr(core, "active_doc", lambda name, attr: None)
    monkeypatch.setattr(core, "doc_alive", lambda obj: False)
    assert app.active_book() is live_book
    assert live_book in app._docs.values()  # 实时句柄并入文档表


def test_active_doc_none_does_not_create(monkeypatch):
    # P0-8：无活动工作簿时 active_book 纯探测返回 None，绝不隐式 Add()
    from offipy import core
    from offipy.excel import ExcelApp

    def _fail_add(*a, **k):
        raise AssertionError("P0-8: active_book 不得隐式 Workbooks.Add()")

    fake_workbooks = type("W", (), {"Add": _fail_add})()
    fake_app = type("F", (), {"ActiveWorkbook": None, "Workbooks": fake_workbooks})()
    app = ExcelApp.__new__(ExcelApp)
    app.app = fake_app
    app._docs = {}
    app._active_id = None
    app._seq = 0
    monkeypatch.setattr(core, "active_doc", lambda name, attr: None)
    monkeypatch.setattr(core, "doc_alive", lambda obj: False)
    assert app.active_book() is None


def test_active_pres_prefers_live(monkeypatch):
    from offipy import core
    from offipy.ppt import PptApp

    live = object()
    app = PptApp.__new__(PptApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    monkeypatch.setattr(core, "active_doc", lambda name, attr: live)
    assert app.active_pres() is live


def test_active_word_doc_prefers_live(monkeypatch):
    from offipy import core
    from offipy.word import WordApp

    live = object()
    app = WordApp.__new__(WordApp)
    app._docs = {}
    app._active_id = None
    app._seq = 0
    monkeypatch.setattr(core, "active_doc", lambda name, attr: live)
    assert app.active_doc() is live
