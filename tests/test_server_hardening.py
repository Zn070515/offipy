"""Batch 3（§4/§5）：超时对齐 / request_id 幂等 / 有界队列与线程 / PID-token 启动锁 / token 权限。

起真实 ThreadingHTTPServer + 假 dispatch（隔离 COM），覆盖：
- client 与 server 的 _CALL_TIMEOUT 单源一致（600）
- request_id 幂等：同 id 重复请求不重执行，命中缓存返回同一响应
- COM 队列满 → /call 503 busy（不无限排队）
- 并发连接超限 → 连接级 503（线程有界）
- PID 文件新格式 {port,pid,token_sha256}：server 落盘 / client 归属验证
- token 落盘 chmod 0o600
"""

import hashlib
import http.client
import json
import os
import queue
import threading
import time

import pytest

from offipy import client, server

TOKEN = "test-token-abc"


def _get(port: int, path: str, token: str | None = None):
    headers = {"Authorization": f"Bearer {token}"} if token else {}
    conn = http.client.HTTPConnection("127.0.0.1", port)
    conn.request("GET", path, headers=headers)
    resp = conn.getresponse()
    data = resp.read().decode("utf-8")
    conn.close()
    return resp.status, json.loads(data) if data else None


def _post(port: int, body: dict, token: str | None = None):
    data = json.dumps(body).encode("utf-8")
    headers = {"Content-Type": "application/json", "X-Offipy-Protocol": "offipy-http/v1"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    conn = http.client.HTTPConnection("127.0.0.1", port)
    conn.request("POST", "/call", body=data, headers=headers)
    resp = conn.getresponse()
    payload = resp.read().decode("utf-8")
    conn.close()
    return resp.status, json.loads(payload) if payload else None


@pytest.fixture
def srv(monkeypatch):
    """真实线程 server + 假 dispatch（隔离 COM），暴露调用记录。"""
    calls = []

    def fake_dispatch(app, op, args, app_name):
        calls.append((app_name, op, args))
        if op == "slow":
            time.sleep(0.2)
            return "slow-done"
        if op == "boom":
            raise ValueError("boom op")
        return {"calls": len(calls)}

    monkeypatch.setattr(server, "get_app", lambda name: object())
    monkeypatch.setattr(server, "dispatch", fake_dispatch)
    monkeypatch.setitem(server._OPS, "ppt", server._OPS["ppt"] | {"slow", "boom"})

    server._TOKEN = TOKEN
    server._ensure_worker()
    srv = server.Server(("127.0.0.1", 0), server.Handler)
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
    yield port, calls
    srv.shutdown()
    srv.server_close()
    server._stop_worker()


def _wait_ready(port: int) -> None:
    for _ in range(100):
        try:
            status, _ = _get(port, "/ping")
            if status == 200:
                return
        except OSError:
            time.sleep(0.05)
    raise RuntimeError("server 未就绪")


# --- 超时对齐（§4） ---


def test_timeout_constant_single_source():
    # client 120 vs server 600 错配修复：两边同名常量、取值一致
    assert client._CALL_TIMEOUT == server._CALL_TIMEOUT == 600


# --- request_id 幂等（§4） ---


def test_request_id_same_id_no_double_exec(srv):
    port, calls = srv
    body = {"app": "ppt", "op": "slow", "request_id": "rid-0001"}
    s1, r1 = _post(port, body, token=TOKEN)
    assert s1 == 200 and r1["data"] == "slow-done"
    s2, r2 = _post(port, body, token=TOKEN)
    assert s2 == 200
    assert r2 == r1  # 命中缓存，返回同一响应
    assert len(calls) == 1  # 只执行一次，第二次不重执行


def test_request_id_different_ids_both_execute(srv):
    port, calls = srv
    _post(port, {"app": "ppt", "op": "slow", "request_id": "rid-a"}, token=TOKEN)
    _post(port, {"app": "ppt", "op": "slow", "request_id": "rid-b"}, token=TOKEN)
    assert len(calls) == 2


def test_request_id_no_id_no_dedupe(srv):
    # 不传 request_id（旧 client）→ 每次照常执行，不误伤
    port, calls = srv
    _post(port, {"app": "ppt", "op": "slow"}, token=TOKEN)
    _post(port, {"app": "ppt", "op": "slow"}, token=TOKEN)
    assert len(calls) == 2


def test_request_id_failure_also_deduped(srv):
    # 失败响应同样缓存：重试不重执行已失败的操作
    port, calls = srv
    body = {"app": "ppt", "op": "boom", "request_id": "rid-err"}
    s1, r1 = _post(port, body, token=TOKEN)
    assert s1 == 500
    s2, r2 = _post(port, body, token=TOKEN)
    assert s2 == 500 and r2 == r1
    assert len(calls) == 1


def test_dedupe_cache_lru_cap(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    monkeypatch.setattr(server, "_REQUEST_ID_MAX", 3)
    for i in range(5):
        server._dedupe_store(f"rid-{i}", {"ok": True, "i": i})
    assert len(server._REQUEST_ID_CACHE) == 3
    assert server._dedupe_hit("rid-0") is None  # 最旧被淘汰
    assert server._dedupe_hit("rid-4") == {"ok": True, "i": 4}


# --- 有界队列（§4）：满 → 503 ---


def test_com_queue_full_returns_503(monkeypatch):
    monkeypatch.setattr(server, "_ensure_worker", lambda: None)
    monkeypatch.setitem(server._OPS, "ppt", server._OPS["ppt"] | {"order"})
    full = queue.Queue(maxsize=1)
    full.put(("ppt", "order", {}, queue.Queue()))  # 永不消费的占位，恒满
    monkeypatch.setattr(server, "_COM_QUEUE", full)
    server._TOKEN = TOKEN
    srv = server.Server(("127.0.0.1", 0), server.Handler)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    port = srv.server_address[1]
    try:
        _wait_ready(port)
        status, body = _post(port, {"app": "ppt", "op": "order"}, token=TOKEN)
        assert status == 503
        assert body["error_code"] == "busy"
        assert "忙" in body["error"]
    finally:
        srv.shutdown()
        srv.server_close()


# --- 有界线程（§4）：并发超限 → 连接级 503 ---


def test_concurrency_limit_503(monkeypatch):
    monkeypatch.setitem(server._OPS, "ppt", server._OPS["ppt"] | {"order"})
    server._TOKEN = TOKEN
    srv = server.Server(("127.0.0.1", 0), server.Handler)
    srv._SLOT = threading.BoundedSemaphore(1)  # 实例级压到 1
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    port = srv.server_address[1]
    try:
        _wait_ready(port)  # 就绪探测时槽位空闲，/ping 200
        srv._SLOT.acquire()  # 占满唯一槽位：后续连接全部 503
        status, body = _post(port, {"app": "ppt", "op": "order"}, token=TOKEN)
        assert status == 503
        assert body["error_code"] == "busy"
    finally:
        srv._SLOT.release()
        srv.shutdown()
        srv.server_close()


def test_max_concurrency_constant():
    assert server._MAX_CONCURRENCY == 16
    assert server.Server._SLOT._initial_value == server._MAX_CONCURRENCY


# --- PID 文件新格式（§4 启动锁） ---


def test_serve_writes_pid_file_json(monkeypatch, tmp_path):
    monkeypatch.setattr(server, "user_data_dir", lambda: tmp_path)
    server._write_pid_file(server.DEFAULT_PORT, "tok-secret")
    data = json.loads((tmp_path / "server.pid").read_text(encoding="utf-8"))
    assert data["port"] == server.DEFAULT_PORT
    assert data["pid"] == os.getpid()
    assert data["token_sha256"] == hashlib.sha256(b"tok-secret").hexdigest()
    assert isinstance(data["started_at"], float)


def test_client_pid_file_matches_json_with_token(monkeypatch, tmp_path):
    token = "client-token"
    pid = 4242
    digest = hashlib.sha256(token.encode()).hexdigest()
    (tmp_path / "server.pid").write_text(
        json.dumps({"port": client.PORT, "pid": pid, "token_sha256": digest}),
        encoding="utf-8",
    )
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    monkeypatch.setattr(client, "_token", lambda: token)
    assert client._pid_file_matches(pid) is True
    assert client._pid_file_matches(pid + 1) is False  # pid 不符
    monkeypatch.setattr(client, "_token", lambda: "other-token")
    assert client._pid_file_matches(pid) is False  # token 不符 → 拒绝


def test_client_pid_file_matches_old_plain(monkeypatch, tmp_path):
    # 旧格式纯数字：退化为仅 pid 比对（向后兼容）
    (tmp_path / "server.pid").write_text("12345", encoding="utf-8")
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    assert client._pid_file_matches(12345) is True
    assert client._pid_file_matches(1) is False


def test_find_server_pid_json_format(monkeypatch, tmp_path):
    (tmp_path / "server.pid").write_text(
        json.dumps({"port": client.PORT, "pid": 777}), encoding="utf-8"
    )
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    monkeypatch.setattr(
        "offipy.client.subprocess.run", lambda *a, **k: type("R", (), {"stdout": ""})()
    )
    assert client._find_server_pid() == 777


# --- token 权限（§4） ---


def test_load_token_chmod_0600(monkeypatch, tmp_path):
    monkeypatch.delenv("OFFIPY_SERVER_TOKEN", raising=False)
    monkeypatch.setattr(server, "user_data_dir", lambda: tmp_path)
    calls = []
    real = os.chmod

    def spy(path, mode):
        calls.append((str(path), mode))
        real(path, mode)

    monkeypatch.setattr(os, "chmod", spy)
    server._load_token(server.DEFAULT_PORT)
    assert (str(tmp_path / "token"), 0o600) in calls
