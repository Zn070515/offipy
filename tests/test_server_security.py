"""server 安全回归（P0-4）：token 鉴权 / status / 请求限制 / op 白名单。

起一个临时端口的真实 HTTPServer 实例，只测安全层——不 dispatch 真实 op，
因此不需要 COM/Office。401 不杀 server 是本轮关键行为（旧 client 连新
server 只报错，不自杀进程）。
"""

import http.client
import json
import socket
import threading
import time

import pytest

from offipy import server

TOKEN = "test-token-abc"


def _get(port: int, path: str, token: str | None = None):
    headers = {"Authorization": f"Bearer {token}"} if token else {}
    conn = http.client.HTTPConnection("127.0.0.1", port)
    conn.request("GET", path, headers=headers)
    resp = conn.getresponse()
    data = resp.read().decode("utf-8")
    conn.close()
    return resp.status, json.loads(data) if data else None


def _post(
    port: int,
    body: dict,
    token: str | None = None,
    ctype="application/json",
    content_length: int | None = None,
    path="/call",
    protocol="offipy-http/v1",
):
    data = json.dumps(body).encode("utf-8")
    headers = {"Content-Type": ctype}
    if protocol is not None:
        headers["X-Offipy-Protocol"] = protocol
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if content_length is not None:
        headers["Content-Length"] = str(content_length)
    conn = http.client.HTTPConnection("127.0.0.1", port)
    conn.request("POST", path, body=data, headers=headers)
    resp = conn.getresponse()
    payload = resp.read().decode("utf-8")
    conn.close()
    return resp.status, json.loads(payload) if payload else None


@pytest.fixture
def srv():
    srv = server.Server(("127.0.0.1", 0), server.Handler)
    server._TOKEN = TOKEN
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    port = srv.server_address[1]
    for _ in range(100):
        try:
            status, _ = _get(port, "/ping")
            if status == 200:
                break
        except OSError:
            time.sleep(0.05)
    yield port
    srv.shutdown()
    srv.server_close()


def test_ping_no_auth_ok(srv):
    status, body = _get(srv, "/ping")
    assert status == 200
    assert body["result"] == "pong"


def test_status_requires_auth(srv):
    status, _ = _get(srv, "/status")
    assert status == 401
    status, body = _get(srv, "/status", token=TOKEN)
    assert status == 200
    result = body["result"]
    for field in ("version", "protocol", "pid", "python", "started_at", "session_id"):
        assert field in result
    assert result["protocol"] == "offipy-http/v1"
    assert result["version"]
    assert result["session_id"]  # P2-3：会话标识随 /status 暴露


def test_status_targets_null_when_uninitialized(srv):
    # P0-7/P0-8：/status 报各 App 目标身份；未初始化（未拉 Office）的 App 报
    # null，绝不因 status 探测而拉起 Office 进程。
    status, body = _get(srv, "/status", token=TOKEN)
    assert status == 200
    targets = body["result"]["targets"]
    assert set(targets) == {"excel", "word", "ppt", "diagram"}
    assert targets["excel"] is None
    assert targets["word"] is None
    assert targets["ppt"] is None
    assert targets["diagram"] is None  # 纯函数 app（diagram）无 COM 根，同样报 null


def test_call_requires_auth(srv):
    status, _ = _post(srv, {"app": "ppt", "op": "quit"})
    assert status == 401
    status, _ = _post(srv, {"app": "ppt", "op": "quit"}, token="wrong-token")
    assert status == 401


def test_call_requires_protocol_header(srv):
    # P2-8：/call 不带 protocol 头 → 400 + ProtocolError
    status, body = _post(srv, {"app": "ppt", "op": "quit"}, token=TOKEN, protocol=None)
    assert status == 400
    assert body["error_code"] == "protocol"


def test_call_rejects_wrong_protocol(srv):
    # P2-8：协议版本不匹配 → 400 + ProtocolError，不静默错位
    status, body = _post(srv, {"app": "ppt", "op": "quit"}, token=TOKEN, protocol="offipy-http/v2")
    assert status == 400
    assert body["error_code"] == "protocol"
    assert "offipy-http/v1" in body["error"]


def test_shutdown_requires_protocol_header(srv):
    # P2-8：/shutdown 同样要求 protocol 头
    status, body = _post(srv, {}, token=TOKEN, path="/shutdown", protocol=None)
    assert status == 400
    assert body["error_code"] == "protocol"


def test_content_type_must_be_json(srv):
    status, _ = _post(srv, {"app": "ppt", "op": "quit"}, token=TOKEN, ctype="text/plain")
    assert status == 415


def test_body_too_large_rejected(srv):
    status, _ = _post(
        srv,
        {"app": "ppt", "op": "quit"},
        token=TOKEN,
        content_length=server._MAX_BODY + 1,
    )
    assert status == 413


def test_unknown_app_rejected(srv):
    status, body = _post(srv, {"app": "bogus", "op": "x"}, token=TOKEN)
    assert status == 400
    assert "未知应用" in body["error"]


def test_op_not_in_whitelist_rejected(srv):
    status, _ = _post(srv, {"app": "ppt", "op": "does_not_exist"}, token=TOKEN)
    assert status == 400
    # 私有/dunder 方法不进白名单
    status, _ = _post(srv, {"app": "ppt", "op": "__class__"}, token=TOKEN)
    assert status == 400


def test_server_survives_bad_auth(srv):
    # 401 不杀 server：坏 token 之后 /ping 与正 token 调用仍可用
    _post(srv, {"app": "ppt", "op": "quit"}, token="wrong")
    status, _ = _get(srv, "/ping")
    assert status == 200
    status, _ = _get(srv, "/status", token=TOKEN)
    assert status == 200


def test_post_non_call_path_404(srv):
    # P0-10：POST 非 /call 路径 → 404，不进入 dispatch
    status, body = _post(srv, {"app": "ppt", "op": "quit"}, token=TOKEN, path="/nope")
    assert status == 404
    assert "not found" in body["error"]


def test_negative_content_length_rejected(srv):
    # P0-10：负 Content-Length → 400（read(负值) 会吞掉整个连接缓冲）
    status, _ = _post(srv, {"app": "ppt", "op": "quit"}, token=TOKEN, content_length=-5)
    assert status == 400


def test_shutdown_requires_auth(srv):
    status, _ = _post(srv, {}, path="/shutdown")
    assert status == 401
    status, _ = _post(srv, {}, token="wrong-token", path="/shutdown")
    assert status == 401


def test_shutdown_ok(srv):
    status, body = _post(srv, {}, token=TOKEN, path="/shutdown")
    assert status == 200
    assert body["ok"] is True


def test_deprecated_op_gets_warning(monkeypatch):
    # P2-9：schema 标 deprecated 的 op，成功/失败响应都带 warning 字段
    from offipy import schema
    from offipy.schema import OpSpec

    orig = schema.spec

    def fake_spec(app, op):
        sp = orig(app, op)
        if sp is not None and app == "word" and op == "new_doc":
            return OpSpec(
                description=sp.description,
                readonly=sp.readonly,
                destructive=sp.destructive,
                deprecated=True,
                returns=sp.returns,
                params=sp.params,
            )
        return sp

    monkeypatch.setattr(schema, "spec", fake_spec)
    ok = server._success_result("word.new_doc", None)
    assert "warning" in ok and "已弃用" in ok["warning"]
    err = server._error_result("word.new_doc", ValueError("x"))
    assert "warning" in err
    plain = server._success_result("word.write_line", None)
    assert "warning" not in plain


def test_encode_reply_ok_normal():
    status, data = server._encode_reply({"ok": True, "result": 3})
    assert status == 200
    assert json.loads(data) == {"ok": True, "result": 3}


def test_encode_reply_caps_oversize(monkeypatch):
    # P0-10：响应超上限 → 500 错误，不向客户端写大 payload
    monkeypatch.setattr(server, "_MAX_RESPONSE", 100)
    status, data = server._encode_reply({"ok": True, "result": "x" * 1000})
    assert status == 500
    payload = json.loads(data)
    assert payload["ok"] is False
    assert "上限" in payload["error"]


# --- round-2：host 限制 / token 生命周期 / 显式白名单 / 惰性 COM ---


def test_validate_host_rejects_non_loopback():
    with pytest.raises(server.ServerStartError):
        server._validate_host("0.0.0.0", allow_remote=False)
    with pytest.raises(server.ServerStartError):
        server._validate_host("192.168.1.5", allow_remote=False)


def test_validate_host_allows_loopback_and_explicit_remote():
    for host in ("127.0.0.1", "localhost", "::1", ""):
        server._validate_host(host, allow_remote=False)  # 不抛
    server._validate_host("0.0.0.0", allow_remote=True)  # 显式放行不抛


def test_warn_if_remote_prints_plaintext_warning(capsys):
    # D6：非回环绑定打印明文传输警告；"" 绑定所有接口同样警告；回环不警告
    server._warn_if_remote("0.0.0.0")
    out = capsys.readouterr().out
    assert "明文" in out and "TLS" in out and "0.0.0.0" in out
    server._warn_if_remote("")
    assert "明文" in capsys.readouterr().out  # "" = INADDR_ANY，同样警告
    server._warn_if_remote("127.0.0.1")
    server._warn_if_remote("localhost")
    server._warn_if_remote("::1")
    assert capsys.readouterr().out == ""  # 回环不警告


def test_load_token_env_first_no_file_write(monkeypatch, tmp_path):
    monkeypatch.setenv("OFFIPY_SERVER_TOKEN", "env-token-xyz")
    monkeypatch.setattr(server, "user_data_dir", lambda: tmp_path)
    assert server._load_token(server.DEFAULT_PORT) == "env-token-xyz"
    assert not (tmp_path / "token").exists()  # env 模式不落盘


def test_load_token_write_failure_raises(monkeypatch, tmp_path):
    # token 文件写不了 → ServerStartError，杜绝「server 假活、client 必 401」
    monkeypatch.delenv("OFFIPY_SERVER_TOKEN", raising=False)
    blocker = tmp_path / "blocker"
    blocker.write_text("x")  # 用文件挡住 user_data_dir，mkdir 必失败
    monkeypatch.setattr(server, "user_data_dir", lambda: blocker)
    with pytest.raises(server.ServerStartError):
        server._load_token(server.DEFAULT_PORT)


def test_session_internal_ops_not_in_whitelist():
    # P1-1 显式注册表：active_pres/active_doc/active_book 一律不暴露
    for app in ("ppt", "word", "excel"):
        assert "active_pres" not in server._OPS.get(app, frozenset())
        assert "active_doc" not in server._OPS.get(app, frozenset())
        assert "active_book" not in server._OPS.get(app, frozenset())
        assert "quit" in server._OPS[app]
        assert "save" in server._OPS[app]


def test_read_ops_in_whitelist():
    # Agent 只读 op 已登记：word read_doc_text / ppt read_slide_texts / excel read_range
    assert "read_doc_text" in server._OPS["word"]
    assert "read_slide_texts" in server._OPS["ppt"]
    assert "read_range" in server._OPS["excel"]


def test_serialize_recurses_dict():
    # read_slide_texts 返回 list[dict]，dict 必须递归序列化而非 str(dict)
    assert server._serialize({"a": [1, 2], "b": {"c": "x"}}) == {"a": [1, 2], "b": {"c": "x"}}
    assert server._serialize([{"index": 1, "title": "t"}]) == [{"index": 1, "title": "t"}]


def test_server_module_lazy_com():
    # 顶层不再裸 import COM：跨平台 `import offipy.server` 不炸
    assert not hasattr(server, "pythoncom")
    assert not hasattr(server, "pywintypes")
    err = server._com_error()
    assert isinstance(err, type) and issubclass(err, BaseException)


def test_slowloris_idle_connection_closed(monkeypatch):
    # H10（slowloris）：半连接超过 socket 读超时即被断开——否则空连接占满
    # 并发槽位拖垮 /call。monkeypatch 缩短超时以便测试。
    monkeypatch.setattr(server.Handler, "timeout", 0.3)
    server._TOKEN = TOKEN
    srv = server.Server(("127.0.0.1", 0), server.Handler)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    port = srv.server_address[1]
    sock = socket.create_connection(("127.0.0.1", port))
    try:
        for _ in range(100):
            try:
                status, _ = _get(port, "/ping")
                if status == 200:
                    break
            except OSError:
                time.sleep(0.05)
        # 只发半截请求行，不发完整请求/头：模拟 slowloris 慢读连接
        sock.sendall(b"POST /call HTTP/1.1\r\n")
        sock.settimeout(2.0)
        start = time.monotonic()
        closed = False
        while time.monotonic() - start < 1.5:
            try:
                data = sock.recv(4096)
            except OSError:
                closed = True  # RST：连接被 server 重置
                break
            if not data:
                closed = True  # EOF：server 已关连接
                break
        assert closed, "server 未在读超时窗口内关闭空闲连接（slowloris 漏洞）"
    finally:
        sock.close()
        srv.shutdown()
        srv.server_close()
