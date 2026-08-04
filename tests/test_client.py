import io
import json
import urllib.error

import pytest

from offipy import client
from offipy.client import call, convert_value, ensure_server, request
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


# --- 握手 / 拉起 ---


def test_ensure_server_returns_when_alive(monkeypatch):
    monkeypatch.setattr("offipy.client._server_ok", lambda: True)
    assert ensure_server() is None


def test_ensure_server_raises_on_timeout(monkeypatch, tmp_path):
    class FakePopen:
        def __init__(self, *a, **k):
            self.pid = 1

    monkeypatch.setattr("offipy.client._server_ok", lambda: False)
    monkeypatch.setattr("offipy.client._ping", lambda: False)
    monkeypatch.setattr("offipy.client.time.sleep", lambda s: None)
    monkeypatch.setattr("offipy.client.subprocess.Popen", FakePopen)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "offipy server" in str(exc.value)


def test_ensure_server_replaces_stale(monkeypatch, tmp_path):
    # 端口上有进程但握手失败（旧 server）→ 杀之并重启
    calls = {"n": 0}

    def fake_ok():
        calls["n"] += 1
        return calls["n"] > 1

    killed = []

    class FakePopen:
        def __init__(self, *a, **k):
            self.pid = 42

    monkeypatch.setattr("offipy.client._server_ok", fake_ok)
    monkeypatch.setattr("offipy.client._ping", lambda: True)
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 987)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    monkeypatch.setattr("offipy.client.subprocess.Popen", FakePopen)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    ensure_server()
    assert killed == [987]


def test_ensure_server_stale_but_unfindable_raises(monkeypatch, tmp_path):
    monkeypatch.setattr("offipy.client._server_ok", lambda: False)
    monkeypatch.setattr("offipy.client._ping", lambda: True)
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: None)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "offipy server stop" in str(exc.value)


def test_server_ok_true_on_matching_status(monkeypatch):
    class _Resp:
        status = 200

        def read(self):
            return b'{"ok": true, "result": {"protocol": "offipy-http/v1"}}'

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(
        "offipy.client._OPENER.open", lambda url, headers=None, timeout=None: _Resp()
    )
    assert client._server_ok() is True


def test_server_ok_false_on_error(monkeypatch):
    monkeypatch.setattr(
        "offipy.client._OPENER.open",
        lambda url, headers=None, timeout=None: (_ for _ in ()).throw(
            urllib.error.URLError("refused")
        ),
    )
    assert client._server_ok() is False


# --- HTTP 错误语义化（P0-1） ---


def test_request_wraps_http_error(monkeypatch):
    def raiser(req, timeout=None):
        body = json.dumps({"ok": False, "error": "unauthorized"}).encode()
        raise urllib.error.HTTPError(req.full_url, 401, "Unauthorized", {}, io.BytesIO(body))

    monkeypatch.setattr("offipy.client._server_ok", lambda: True)
    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    with pytest.raises(RemoteCallError) as exc:
        request("ppt", "save")
    assert "unauthorized" in str(exc.value)


def test_request_wraps_urlerror(monkeypatch):
    def raiser(req, timeout=None):
        raise urllib.error.URLError("connection refused")

    monkeypatch.setattr("offipy.client._server_ok", lambda: True)
    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    with pytest.raises(RemoteCallError) as exc:
        request("ppt", "save")
    assert "连接失败" in str(exc.value)


def test_request_wraps_bad_json(monkeypatch):
    class _Resp:
        status = 200

        def read(self):
            return b"not json{{"

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr("offipy.client._server_ok", lambda: True)
    monkeypatch.setattr("offipy.client._OPENER.open", lambda req, timeout=None: _Resp())
    with pytest.raises(RemoteCallError) as exc:
        request("ppt", "save")
    assert "非 JSON" in str(exc.value)


# --- pid 定位 / stop ---


def test_find_server_pid_from_file(monkeypatch, tmp_path):
    # netstat 查不到监听 → 兜底读 pid 文件
    class _R:
        stdout = ""

    (tmp_path / "server.pid").write_text("12345", encoding="utf-8")
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    monkeypatch.setattr("offipy.client.subprocess.run", lambda *a, **k: _R())
    assert client._find_server_pid() == 12345


def test_find_server_pid_netstat_fallback(monkeypatch, tmp_path):
    class _R:
        stdout = "TCP    127.0.0.1:8890    0.0.0.0:0    LISTENING    4321\n"

    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    monkeypatch.setattr("offipy.client.subprocess.run", lambda *a, **k: _R())
    assert client._find_server_pid() == 4321


def test_stop_server_kills(monkeypatch):
    killed = []
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 12345)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    assert client.stop_server() is True
    assert killed == [12345]


def test_stop_server_no_process(monkeypatch):
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: None)
    assert client.stop_server() is False
