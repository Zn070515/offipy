import io
import json
import os
import urllib.error

import pytest

from offipy import client
from offipy.client import call, ensure_server, request
from offipy.exceptions import RemoteCallError, ServerStartError


class _Ctx:
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


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
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")
    assert ensure_server() is None


def test_ensure_server_raises_on_timeout(monkeypatch, tmp_path):
    class FakePopen:
        def __init__(self, *a, **k):
            self.pid = 1

    monkeypatch.setattr("offipy.client._probe", lambda: "down")
    monkeypatch.setattr("offipy.client.time.sleep", lambda s: None)
    monkeypatch.setattr("offipy.client.subprocess.Popen", FakePopen)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "offipy server" in str(exc.value)


def test_ensure_server_auth_fail_never_kills(monkeypatch, tmp_path):
    # P0-2：token 失配 → 抛错，绝不强杀进程
    killed = []
    monkeypatch.setattr("offipy.client._probe", lambda: "auth_fail")
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 987)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "token 不匹配" in str(exc.value)
    assert killed == []


def test_ensure_server_mismatch_owned_kills_and_restarts(monkeypatch, tmp_path):
    # 我们的旧版 server（pid 文件证明归属）→ 可强杀重启
    calls = {"n": 0}

    def fake_probe():
        calls["n"] += 1
        return "mismatch" if calls["n"] == 1 else "ok"

    killed = []

    class FakePopen:
        def __init__(self, *a, **k):
            self.pid = 42

    monkeypatch.setattr("offipy.client._probe", fake_probe)
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 987)
    monkeypatch.setattr("offipy.client._pid_file_matches", lambda pid: True)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    monkeypatch.setattr("offipy.client.subprocess.Popen", FakePopen)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    ensure_server()
    assert killed == [987]


def test_ensure_server_restarts_on_version_skew_owned(monkeypatch, tmp_path):
    # #34：版本偏斜 + pid 文件证明归属 → 强杀重启（走 mismatch 分支）
    calls = {"n": 0}

    def fake_probe():
        calls["n"] += 1
        return "mismatch" if calls["n"] == 1 else "ok"

    killed = []

    class FakePopen:
        def __init__(self, *a, **k):
            self.pid = 42

    monkeypatch.setattr("offipy.client._probe", fake_probe)
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 987)
    monkeypatch.setattr("offipy.client._pid_file_matches", lambda pid: True)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    monkeypatch.setattr("offipy.client.subprocess.Popen", FakePopen)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    ensure_server()
    assert killed == [987]


def test_ensure_server_mismatch_unknown_refuses(monkeypatch, tmp_path):
    # P0-1：非 offipy 进程占端口且无法证明归属 → 拒绝强杀
    killed = []
    monkeypatch.setattr("offipy.client._probe", lambda: "mismatch")
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 987)
    monkeypatch.setattr("offipy.client._pid_file_matches", lambda pid: False)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "拒绝强杀" in str(exc.value)
    assert killed == []


def test_ensure_server_mismatch_unfindable_raises(monkeypatch, tmp_path):
    monkeypatch.setattr("offipy.client._probe", lambda: "mismatch")
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: None)
    monkeypatch.setattr("offipy.client.user_data_dir", lambda: tmp_path)
    with pytest.raises(ServerStartError) as exc:
        ensure_server()
    assert "offipy server stop" in str(exc.value)


# --- _probe 四态 ---


def test_probe_ok_on_matching_status(monkeypatch):
    from offipy import __version__

    class _Resp:
        status = 200

        def read(self):
            return json.dumps(
                {"ok": True, "result": {"protocol": "offipy-http/v1", "version": __version__}}
            ).encode()

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr("offipy.client._OPENER.open", lambda url, timeout=None: _Resp())
    assert client._probe() == "ok"


def test_probe_mismatch_on_version_skew(monkeypatch):
    # #34：协议匹配但版本不一致 → mismatch（stale server），ensure_server 据此重启
    class _Resp:
        status = 200

        def read(self):
            return b'{"ok": true, "result": {"protocol": "offipy-http/v1", "version": "0.0.0-fake"}}'

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr("offipy.client._OPENER.open", lambda url, timeout=None: _Resp())
    assert client._probe() == "mismatch"


def test_probe_auth_fail_on_401(monkeypatch):
    def raiser(url, timeout=None):
        raise urllib.error.HTTPError(
            "http://127.0.0.1:8890/status", 401, "Unauthorized", {}, io.BytesIO(b"")
        )

    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    assert client._probe() == "auth_fail"


def test_probe_mismatch_on_wrong_protocol(monkeypatch):
    class _Resp:
        status = 200

        def read(self):
            return b'{"ok": true, "result": {"protocol": "v0"}}'

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr("offipy.client._OPENER.open", lambda url, timeout=None: _Resp())
    assert client._probe() == "mismatch"


def test_probe_down_on_urlerror(monkeypatch):
    monkeypatch.setattr(
        "offipy.client._OPENER.open",
        lambda url, timeout=None: (_ for _ in ()).throw(urllib.error.URLError("refused")),
    )
    assert client._probe() == "down"


def test_server_ok_true_on_matching_status(monkeypatch):
    from offipy import __version__

    class _Resp:
        status = 200

        def read(self):
            return json.dumps(
                {"ok": True, "result": {"protocol": "offipy-http/v1", "version": __version__}}
            ).encode()

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


def test_server_status_returns_dict_on_version_skew(monkeypatch):
    # #34：版本偏斜但协议匹配 → 返回可读 dict（含 version），不因偏斜吞成 None
    class _Resp:
        status = 200

        def read(self):
            return b'{"ok": true, "result": {"protocol": "offipy-http/v1", "version": "0.0.0-fake"}}'

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr("offipy.client._probe", lambda: "mismatch")
    monkeypatch.setattr("offipy.client._OPENER.open", lambda url, timeout=None: _Resp())
    st = client.server_status()
    assert st is not None
    assert st["version"] == "0.0.0-fake"


def test_server_status_none_on_protocol_mismatch(monkeypatch):
    # #34 补充：非 offipy 进程（协议失配）→ None，不把异质进程暴露成可读 server 状态
    class _Resp:
        status = 200

        def read(self):
            return b'{"ok": true, "result": {"protocol": "v0", "version": "x"}}'

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    monkeypatch.setattr("offipy.client._probe", lambda: "mismatch")
    monkeypatch.setattr("offipy.client._OPENER.open", lambda url, timeout=None: _Resp())
    assert client.server_status() is None


# --- HTTP 错误语义化（P0-1） ---


def test_request_wraps_http_error(monkeypatch):
    def raiser(req, timeout=None):
        body = json.dumps({"ok": False, "error": "unauthorized"}).encode()
        raise urllib.error.HTTPError(req.full_url, 401, "Unauthorized", {}, io.BytesIO(body))

    monkeypatch.setattr("offipy.client._probe", lambda: "ok")
    monkeypatch.setattr("offipy.client._OPENER.open", raiser)
    with pytest.raises(RemoteCallError) as exc:
        request("ppt", "save")
    assert "unauthorized" in str(exc.value)


def test_request_wraps_urlerror(monkeypatch):
    def raiser(req, timeout=None):
        raise urllib.error.URLError("connection refused")

    monkeypatch.setattr("offipy.client._probe", lambda: "ok")
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

    monkeypatch.setattr("offipy.client._probe", lambda: "ok")
    monkeypatch.setattr("offipy.client._OPENER.open", lambda req, timeout=None: _Resp())
    with pytest.raises(RemoteCallError) as exc:
        request("ppt", "save")
    assert "非 JSON" in str(exc.value)


# --- P1-1 端口感知 token / P1-2 expected_target.path 绝对化 ---


def test_port_from_url(monkeypatch):
    monkeypatch.setattr("offipy.client.port", lambda: 8890)
    assert client._port_from_url("http://127.0.0.1:8901") == 8901
    assert client._port_from_url("http://127.0.0.1:8890") == 8890
    assert client._port_from_url("https://example.com/x") == 443
    assert client._port_from_url("http://example.com/x") == 80
    assert client._port_from_url(None) == 8890


def test_token_reads_per_port_file(monkeypatch, tmp_path):
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    monkeypatch.delenv("OFFIPY_SERVER_TOKEN", raising=False)
    (tmp_path / "token-8901").write_text("port-token", encoding="utf-8")
    (tmp_path / "token").write_text("default-token", encoding="utf-8")
    assert client._token(8901) == "port-token"
    assert client._token(client.PORT) == "default-token"
    assert client._token(9999) is None


def test_request_uses_per_port_token_and_absolutizes_paths(monkeypatch, tmp_path):
    captured = {}

    class _Resp:
        status = 200

        def read(self):
            return b'{"ok": true, "result": null}'

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_open(req, timeout=None):
        captured["auth"] = req.get_header("Authorization")
        captured["url"] = req.full_url
        captured["data"] = json.loads(req.data.decode("utf-8"))
        return _Resp()

    monkeypatch.setattr("offipy.client._probe", lambda: "ok")
    monkeypatch.setattr("offipy.client._OPENER.open", fake_open)
    monkeypatch.delenv("OFFIPY_SERVER_TOKEN", raising=False)
    (tmp_path / "token-8901").write_text("p8901", encoding="utf-8")
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)

    request(
        "ppt",
        "export_slides",
        base_url="http://127.0.0.1:8901",
        out_dir="rel/out",
        expected_target={"doc_id": "p1", "path": "rel/notes.docx"},
    )
    args = captured["data"]["args"]
    assert args["out_dir"] == os.path.abspath("rel/out")
    assert args["expected_target"]["path"] == os.path.abspath("rel/notes.docx")
    assert args["expected_target"]["doc_id"] == "p1"  # 其余字段保留
    assert captured["url"] == "http://127.0.0.1:8901/call"
    assert captured["auth"] == "Bearer p8901"  # P1-1：token 按 base_url 端口取


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


def test_stop_server_ok_uses_shutdown(monkeypatch):
    # 可鉴权 → 走 /shutdown 优雅停机，确认进程退出后返回
    called = []
    states = iter(["ok", "down"])
    monkeypatch.setattr("offipy.client._probe", lambda: next(states))
    monkeypatch.setattr(
        "offipy.client._OPENER.open",
        lambda req, timeout=None: called.append(req.full_url) or _Ctx(),
    )
    monkeypatch.setattr("offipy.client.time.sleep", lambda s: None)
    assert client.stop_server() == "server 已停止"
    assert called == ["http://127.0.0.1:8890/shutdown"]


def test_stop_server_ok_old_server_fallback_kill(monkeypatch):
    # 旧版 server 无 /shutdown：等待超时后按 pid 文件归属回退强杀
    called = []
    killed = []
    monkeypatch.setattr("offipy.client._probe", lambda: "ok")  # 永不 down
    monkeypatch.setattr(
        "offipy.client._OPENER.open",
        lambda req, timeout=None: called.append(req.full_url) or _Ctx(),
    )
    monkeypatch.setattr("offipy.client.time.sleep", lambda s: None)
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 12345)
    monkeypatch.setattr("offipy.client._pid_file_matches", lambda pid: True)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    assert client.stop_server() == "server 已停止"
    assert called == ["http://127.0.0.1:8890/shutdown"]
    assert killed == [12345]


def test_stop_server_auth_fail_refuses(monkeypatch):
    killed = []
    monkeypatch.setattr("offipy.client._probe", lambda: "auth_fail")
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    assert "token 不匹配" in client.stop_server()
    assert killed == []


def test_stop_server_mismatch_owned_kills(monkeypatch):
    killed = []
    monkeypatch.setattr("offipy.client._probe", lambda: "mismatch")
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 12345)
    monkeypatch.setattr("offipy.client._pid_file_matches", lambda pid: True)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    assert client.stop_server() == "server 已停止"
    assert killed == [12345]


def test_stop_server_mismatch_unknown_refuses(monkeypatch):
    killed = []
    monkeypatch.setattr("offipy.client._probe", lambda: "mismatch")
    monkeypatch.setattr("offipy.client._find_server_pid", lambda: 12345)
    monkeypatch.setattr("offipy.client._pid_file_matches", lambda pid: False)
    monkeypatch.setattr("offipy.client._kill_pid", lambda pid: killed.append(pid))
    assert "拒绝强杀" in client.stop_server()
    assert killed == []


def test_stop_server_not_running(monkeypatch):
    monkeypatch.setattr("offipy.client._probe", lambda: "down")
    assert client.stop_server() == "server 未在运行"
