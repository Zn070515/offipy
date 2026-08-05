"""P0-4 会话模型：Remote*（client→server）vs direct（本地直连 COM）拆分的纯逻辑测试。

用 mock 断言：
- RemoteExcel/Word/Ppt 构造先 ensure_server（缺省本地 8890），方法调用走 client.call
  （URL/参数/透传正确）
- 私有属性不触发 HTTP（__getattr__ 只代理非下划线名）
- direct 模块与 __init__ 导出一致：direct.Excel is offipy.Excel；Remote 类导出齐
- base_url 显式给出时不触发本地 ensure_server（指向其他 server）
"""

import pytest

import offipy
from offipy import api


def test_remote_excel_ensure_server_and_call(monkeypatch):
    captured = {}

    def fake_call(app, op, base_url=None, request_id=None, **kw):
        captured["app"] = app
        captured["op"] = op
        captured["base_url"] = base_url
        captured["request_id"] = request_id
        captured["kw"] = kw
        return "ok"

    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: "server-up")
    monkeypatch.setattr("offipy.api.client.call", fake_call)
    with api.RemoteExcel() as x:
        assert x.new_book() == "ok"
        x.set_cell(1, "A1", 5, doc_id="book1")
    assert captured == {
        "app": "excel",
        "op": "set_cell",
        "base_url": None,
        "request_id": None,
        "kw": {"sheet": 1, "cell": "A1", "value": 5, "doc_id": "book1"},
    }


def test_remote_request_id_passthrough(monkeypatch):
    # P1-4：Remote 方法额外暴露 request_id（幂等标识），显式给出时透传 client.call
    captured = {}

    def fake_call(app, op, base_url=None, request_id=None, **kw):
        captured["request_id"] = request_id
        return "ok"

    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: "server-up")
    monkeypatch.setattr("offipy.api.client.call", fake_call)
    with api.RemoteWord() as w:
        w.write("hi", request_id="rid-abc")
    assert captured["request_id"] == "rid-abc"


def test_remote_word_ppt_call(monkeypatch):
    calls = []
    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: None)
    monkeypatch.setattr("offipy.api.client.call", lambda *a, **k: calls.append((a, k)))
    with api.RemoteWord() as w:
        w.write_line("hi")
    with api.RemotePpt() as p:
        p.add_slide()
    assert calls[0][0] == ("word", "write_line")
    assert calls[0][1]["base_url"] is None
    assert calls[1][0] == ("ppt", "add_slide")


def test_remote_quit_goes_through_client(monkeypatch):
    calls = []
    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: None)
    monkeypatch.setattr("offipy.api.client.call", lambda *a, **k: calls.append((a, k)) or "bye")
    assert api.RemoteExcel().quit() == "bye"
    assert calls[-1][0] == ("excel", "quit")


def test_remote_private_attr_no_call(monkeypatch):
    calls = []
    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: None)
    monkeypatch.setattr("offipy.api.client.call", lambda *a, **k: calls.append((a, k)))
    x = api.RemoteExcel()
    with pytest.raises(AttributeError):
        _ = x._private  # 私有名不走 HTTP，直接 AttributeError
    assert calls == []


def test_remote_destructive_transport_params(monkeypatch):
    # P0-1/P0-3：破坏性 op 经 Remote facade 透传 follow_active / expected_target
    captured = {}

    def fake_call(app, op, base_url=None, **kw):
        captured.update(kw)
        return "ok"

    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: None)
    monkeypatch.setattr("offipy.api.client.call", fake_call)
    with api.RemoteExcel() as x:
        x.set_cell(1, "A1", 5, follow_active=True)
    assert captured["follow_active"] is True
    assert "expected_target" not in captured  # 未给 → 不下发
    captured.clear()
    with api.RemoteExcel() as x:
        x.set_cell(1, "A1", 5, expected_target={"doc_id": "book1"})
    assert captured["expected_target"] == {"doc_id": "book1"}
    assert "follow_active" not in captured


def test_remote_explicit_base_url_skips_local_ensure(monkeypatch):
    # base_url 指向其他 server：不再探测/拉起本地 8890
    ensured = []

    def fake_call(app, op, base_url=None, **kw):
        return "ok"

    monkeypatch.setattr("offipy.api.client.ensure_server", lambda: ensured.append(1))
    monkeypatch.setattr("offipy.api.client.call", fake_call)
    with api.RemoteExcel(base_url="http://127.0.0.1:9999") as x:
        assert x.new_book() == "ok"
    assert ensured == []  # 未触发本地 ensure_server


def test_direct_reexports_local_facades():
    from offipy import direct

    assert direct.Excel is offipy.Excel
    assert direct.Word is offipy.Word
    assert direct.Ppt is offipy.Ppt


def test_remote_classes_and_direct_exported():
    assert hasattr(offipy, "RemoteExcel")
    assert hasattr(offipy, "RemoteWord")
    assert hasattr(offipy, "RemotePpt")
    assert offipy.direct.Excel is offipy.Excel
