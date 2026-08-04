import pytest

from offipy.client import call, convert_value, ensure_server
from offipy.exceptions import RemoteCallError, ServerStartError


def test_convert_value_bool():
    assert convert_value("true") is True
    assert convert_value("false") is False


def test_convert_value_number():
    assert convert_value("42") == 42
    assert convert_value("3.14") == 3.14


def test_convert_value_none():
    assert convert_value("none") is None


def test_convert_value_case_insensitive():
    assert convert_value("TRUE") is True
    assert convert_value("False") is False
    assert convert_value("None") is None
    assert convert_value(" null ") is None  # 去空白 + null 别名


def test_convert_value_str():
    assert convert_value("hello") == "hello"


def test_call_returns_result_on_ok(monkeypatch):
    monkeypatch.setattr("offipy.client.request", lambda app, op, **a: {"ok": True, "result": 3})
    assert call("ppt", "add_slide") == 3


def test_call_raises_remote_call_error_on_fail(monkeypatch):
    monkeypatch.setattr(
        "offipy.client.request",
        lambda app, op, **a: {"ok": False, "error": "boom", "trace": ["at f()"]},
    )
    with pytest.raises(RemoteCallError) as exc:
        call("ppt", "add_slide")
    assert "boom" in str(exc.value)
    assert "f()" in str(exc.value)


def test_ensure_server_returns_when_alive(monkeypatch):
    monkeypatch.setattr("offipy.client._ping", lambda: True)
    assert ensure_server() is None


def test_ensure_server_raises_on_timeout(monkeypatch, tmp_path):
    class FakePopen:
        def __init__(self, *a, **k):
            pass

    monkeypatch.setattr("offipy.client._ping", lambda: False)
    monkeypatch.setattr("offipy.client.time.sleep", lambda s: None)
    monkeypatch.setattr("offipy.client.subprocess.Popen", FakePopen)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "offipy server" in str(exc.value)
