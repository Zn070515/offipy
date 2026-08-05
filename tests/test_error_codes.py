"""P1-4 error_code 往返 + guard_com 包装 + dispatch 断连重试。

覆盖：
- client.request 把 HTTP 错误 + error_code 映射回领域异常
- client.call ok:false 按 error_code 抛对应领域异常；成功返回 data（result 回退）
- _success_result / _error_result 契约（resource_id / hresult）
- _resource_id 格式 "app:kind:name"
- guard_com：COM 错误包装成 ComOperationError（保留 hresult），私有方法不包装
- dispatch：断连 HRESULT 重建重试；非断连 COM 错误原样上抛；quit 不反拉起
"""

import io
import json
import urllib.error

import pytest

from offipy import _comguard, client, server
from offipy.exceptions import (
    ComOperationError,
    FileConflictError,
    InvalidArgumentError,
    TargetNotFoundError,
)
from offipy.result import OperationResult

# --- client：HTTP 错误 + error_code → 领域异常 ---


def _http_error(req, status, code, hresult=None):
    body = {"ok": False, "error": "boom", "error_code": code}
    if hresult is not None:
        body["hresult"] = hresult
    data = json.dumps(body).encode()
    return urllib.error.HTTPError(req.full_url, status, "Error", {}, io.BytesIO(data))


@pytest.mark.parametrize(
    ("code", "exc_cls"),
    [
        ("invalid_argument", InvalidArgumentError),
        ("target_not_found", TargetNotFoundError),
        ("file_conflict", FileConflictError),
        ("com_operation", ComOperationError),
    ],
)
def test_request_maps_http_error_code(monkeypatch, code, exc_cls):
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")

    def raiser(req, timeout=None):
        raise _http_error(req, 500, code)

    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    with pytest.raises(exc_cls):
        client.request("excel", "set_cell", sheet=1, cell="A1", value=1)


def test_request_unknown_code_falls_back_to_remote(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")

    def raiser(req, timeout=None):
        raise _http_error(req, 500, "some_future_code")

    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    from offipy.exceptions import RemoteCallError

    with pytest.raises(RemoteCallError):
        client.request("excel", "set_cell", sheet=1, cell="A1", value=1)


# --- 契约4/5：request 应用层失败抛异常（非返回 dict）+ ComOperationError 透传 hresult ---


def test_request_com_error_preserves_hresult(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")

    def raiser(req, timeout=None):
        raise _http_error(req, 500, "com_operation", hresult="0x80010108")

    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    with pytest.raises(ComOperationError) as exc:
        client.request("excel", "read_range", sheet=1, range_addr="A1")
    assert exc.value.hresult == 0x80010108


def test_request_com_error_without_hresult_is_none(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")

    def raiser(req, timeout=None):
        raise _http_error(req, 500, "com_operation")

    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    with pytest.raises(ComOperationError) as exc:
        client.request("excel", "read_range", sheet=1, range_addr="A1")
    assert exc.value.hresult is None


def test_call_com_error_preserves_hresult(monkeypatch):
    monkeypatch.setattr(
        "offipy.client.request",
        lambda app, op, **a: {
            "ok": False,
            "error": "COM 失败",
            "error_code": "com_operation",
            "hresult": "0x80004005",
        },
    )
    with pytest.raises(ComOperationError) as exc:
        client.call("excel", "set_cell", sheet=1, cell="A1", value=1)
    assert exc.value.hresult == 0x80004005


def test_call_maps_error_code(monkeypatch):
    monkeypatch.setattr(
        "offipy.client.request",
        lambda app, op, **a: {
            "ok": False,
            "error": "没有打开的工作簿",
            "error_code": "target_not_found",
        },
    )
    with pytest.raises(TargetNotFoundError):
        client.call("excel", "get_target")


def test_call_returns_data_on_ok(monkeypatch):
    # read_slide_texts 现为按页 per-shape（0.10 契约），必须传 slide_idx；
    # mock 数据用真实 SlideTextRecord 而非 0.9 摘要 dict。
    record = {
        "shape_id": 2,
        "name": "Title 1",
        "text": "t",
        "left": 36.0,
        "top": 21.6,
        "width": 648.0,
        "height": 90.0,
        "coordinate_space": "slide",
        "coordinate_unit": "pt",
        "is_placeholder": True,
        "placeholder_type": 1,
        "placeholder_type_name": "title",
        "parent_shape_id": None,
        "group_path": [],
    }
    monkeypatch.setattr(
        "offipy.client.request",
        lambda app, op, **a: {"ok": True, "data": [record]},
    )
    assert client.call("ppt", "read_slide_texts", slide_idx=1) == [record]


# --- OperationResult 契约（P1-3） ---


def test_operation_result_to_dict():
    r = OperationResult(
        ok=True, operation="excel.new_book", resource_id="excel:book:Book1", message="ok", data=3
    )
    assert r.to_dict() == {
        "ok": True,
        "operation": "excel.new_book",
        "resource_id": "excel:book:Book1",
        "message": "ok",
        "data": 3,
    }


def test_operation_result_failure_adds_error():
    r = OperationResult(ok=False, operation="x.y", resource_id=None, message="boom")
    d = r.to_dict()
    assert d["ok"] is False
    assert d["error"] == "boom"
    assert "data" in d and d["data"] is None


# --- server 结果归一（P1-3） ---


def test_success_result_shape(monkeypatch):
    monkeypatch.setattr(server, "_resource_id", lambda app, doc_id=None: "excel:book:Book1")
    res = server._success_result("excel.new_book", 3)
    assert res["ok"] is True
    assert res["operation"] == "excel.new_book"
    assert res["resource_id"] == "excel:book:Book1"
    assert res["message"] == "ok"
    assert res["data"] == 3
    assert res["result"] == 3  # 兼容别名


def test_success_result_serializes_com_to_none():
    class _FakeCom:
        _oleobj_ = object()

    res = server._success_result("excel.new_book", _FakeCom())
    assert res["data"] is None
    assert res["result"] is None


def test_error_result_code_and_hresult():
    e = ComOperationError("com fail", hresult=0x80010111)
    res = server._error_result("excel.set_cell", e)
    assert res["ok"] is False
    assert res["operation"] == "excel.set_cell"
    assert res["resource_id"] is None
    assert res["error_code"] == "com_operation"
    assert res["hresult"] == "0x80010111"
    assert res["trace"]


def test_error_result_plain_exception_internal():
    res = server._error_result("ppt.add_slide", ValueError("x"))
    assert res["error_code"] == "internal"


def test_resource_id_format(monkeypatch):
    class FakeApp:
        def get_target(self):
            return {"app": "excel", "doc_id": "book1", "name": "Book1", "path": "C:/tmp/book.xlsx"}

    class NoTarget:
        def get_target(self):
            return None

    monkeypatch.setitem(server._APPS, "excel", FakeApp())
    monkeypatch.setitem(server._APPS, "word", NoTarget())
    assert server._resource_id("excel") == "excel:book:book1"  # P0-6：用 doc_id 而非 name
    assert server._resource_id("excel", "book1") == "excel:book:book1"  # 显式 doc_id 优先
    assert server._resource_id("ppt") is None  # 未初始化的 App → None
    assert server._resource_id("word") is None  # 有 App 但无目标 → None


# --- guard_com 包装（P1-4） ---


def test_guard_com_wraps_com_error(monkeypatch):
    class FakeComError(Exception):
        hresult = 0x80004005

    monkeypatch.setattr(_comguard, "_COM_ERROR", FakeComError)

    class FakeApp:
        def boom(self):
            raise FakeComError("com blew up")

        def _private(self):
            raise FakeComError("should stay raw")

    _comguard.guard_com(FakeApp)
    with pytest.raises(ComOperationError) as exc:
        FakeApp().boom()
    assert exc.value.hresult == 0x80004005
    assert exc.value.code == "com_operation"
    # 私有方法不包装：原样抛
    with pytest.raises(FakeComError):
        FakeApp()._private()


def test_guard_com_passes_through_non_com_errors(monkeypatch):
    class NopeError(Exception):
        pass

    monkeypatch.setattr(_comguard, "_COM_ERROR", NopeError)

    class FakeApp:
        def nope(self):
            raise ValueError("plain")

    _comguard.guard_com(FakeApp)
    with pytest.raises(ValueError):
        FakeApp().nope()


def test_guard_com_idempotent(monkeypatch):
    class FakeComError(Exception):
        pass

    monkeypatch.setattr(_comguard, "_COM_ERROR", FakeComError)

    class FakeApp:
        def boom(self):
            raise FakeComError("x")

    _comguard.guard_com(FakeApp)
    _comguard.guard_com(FakeApp)  # 二次包装不叠加
    with pytest.raises(ComOperationError):
        FakeApp().boom()


# --- dispatch 断连重试（P1-4） ---


def test_dispatch_retries_on_disconnect(monkeypatch):
    calls = {"n": 0}

    class FakeApp:
        def op(self):
            calls["n"] += 1
            if calls["n"] == 1:
                raise ComOperationError("disconnected", hresult=0x80010111)
            return "recovered"

    monkeypatch.setattr(server, "_alive", lambda a: True)
    assert server.dispatch(FakeApp(), "op", {}, "x") == "recovered"
    assert calls["n"] == 2


def test_dispatch_non_disconnect_com_error_propagates(monkeypatch):
    class FakeApp:
        def op(self):
            raise ComOperationError("boom", hresult=0x80004005)

    monkeypatch.setattr(server, "_alive", lambda a: True)
    with pytest.raises(ComOperationError) as exc:
        server.dispatch(FakeApp(), "op", {}, "x")
    assert exc.value.hresult == 0x80004005


def test_dispatch_quit_skips_rebuild_when_dead(monkeypatch):
    rebuild_calls = []

    class FakeApp:
        def quit(self):
            return None

    monkeypatch.setattr(server, "_alive", lambda a: False)
    monkeypatch.setattr(server, "_rebuild", lambda app: rebuild_calls.append(app) or app)
    assert server.dispatch(FakeApp(), "quit", {}, "x") is None
    assert rebuild_calls == []  # quit 目标就是退出，不反拉起新实例
