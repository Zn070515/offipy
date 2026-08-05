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
import sys
import threading
import time
import types

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
    server._REQUEST_ID_CACHE.clear()  # 幂等缓存是模块级：每测试隔离，防跨测试串味
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
    assert r1["request_id"] == "rid-0001"  # 幂等回显（P0-2 方案 A）
    assert "cached" not in r1  # 首次执行不是缓存命中
    s2, r2 = _post(port, body, token=TOKEN)
    assert s2 == 200
    assert r2["data"] == "slow-done"  # 命中缓存，返回同一结果
    assert r2["cached"] is True  # 缓存命中标注
    assert r2["request_id"] == "rid-0001"
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


def test_call_args_must_be_dict(srv):
    # §11：args 非 dict（list/str）→ 400 invalid_argument，不碰 dispatch
    port, calls = srv
    status, body = _post(port, {"app": "ppt", "op": "slow", "args": [1, 2]}, token=TOKEN)
    assert status == 400
    assert body["error_code"] == "invalid_argument"
    assert "args" in body["error"]
    assert calls == []  # worker 没被调用（fail-fast 在 handler 线程）


def test_call_args_string_rejected(srv):
    port, calls = srv
    status, body = _post(port, {"app": "ppt", "op": "slow", "args": "A1"}, token=TOKEN)
    assert status == 400
    assert body["error_code"] == "invalid_argument"
    assert calls == []


def test_request_id_failure_also_deduped(srv):
    # 失败响应同样缓存：重试不重执行已失败的操作
    port, calls = srv
    body = {"app": "ppt", "op": "boom", "request_id": "rid-err"}
    s1, r1 = _post(port, body, token=TOKEN)
    assert s1 == 500
    assert r1["request_id"] == "rid-err"
    s2, r2 = _post(port, body, token=TOKEN)
    assert s2 == 500
    assert r2["cached"] is True  # 失败结果同样被缓存回放
    assert r2["error_code"] == r1["error_code"] == "internal"
    assert len(calls) == 1


# --- P0-2 方案 A：payload hash 绑定 / in-flight 合并 / LRU ---


def test_request_id_same_id_different_payload_rejected(srv):
    # 同 request_id 不同 payload → 400 invalid_argument，不静默返回旧结果
    port, calls = srv
    base = {"app": "ppt", "op": "slow", "request_id": "rid-dup"}
    s1, r1 = _post(port, {**base, "args": {"a": 1}}, token=TOKEN)
    assert s1 == 200
    s2, r2 = _post(port, {**base, "args": {"a": 2}}, token=TOKEN)
    assert s2 == 400
    assert r2["error_code"] == "invalid_argument"
    assert "不同 payload" in r2["error"]
    assert len(calls) == 1  # 第二次被拒，未执行


def test_request_id_concurrent_same_id_merges(srv):
    # 并发同 ID → 合并执行一次，两个调用方都拿到结果
    port, calls = srv
    body = {"app": "ppt", "op": "slow", "request_id": "rid-merge"}
    results = []
    lock = threading.Lock()

    def _post_in_thread():
        s, r = _post(port, body, token=TOKEN)
        with lock:
            results.append((s, r))

    ts = [threading.Thread(target=_post_in_thread) for _ in range(2)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    assert len(results) == 2
    for s, r in results:
        assert s == 200 and r["data"] == "slow-done"
    assert len(calls) == 1  # 合并：只执行一次


def test_request_id_timeout_then_retry_no_double_exec(srv, monkeypatch):
    # owner 超时(504)后同 id 重试：合并等 worker 完成，不重执行（绝无双写）
    port, calls = srv
    monkeypatch.setattr(server, "_CALL_TIMEOUT", 0.3)
    gate = threading.Event()
    entered = threading.Event()

    def gated_dispatch(app, op, args, app_name):
        calls.append((app_name, op, args))
        entered.set()
        gate.wait(5)  # 模拟慢 op：阻塞直到测试放行
        return "slow-done"

    monkeypatch.setattr(server, "dispatch", gated_dispatch)
    calls.clear()
    body = {"app": "ppt", "op": "slow", "request_id": "rid-gated"}

    s1, r1 = _post(port, body, token=TOKEN)
    assert s1 == 504  # owner 在 _CALL_TIMEOUT 内没等到 worker → 超时
    assert entered.wait(2)  # worker 已进入（op 只执行了一次）

    gate.set()  # 放行 worker
    s2, r2 = _post(port, body, token=TOKEN)
    assert s2 == 200 and r2["data"] == "slow-done"
    assert len(calls) == 1  # 重试不重执行


def test_claim_inflight_merges_after_owner_timeout(monkeypatch):
    # 状态机：owner 超时后 entry 留 inflight，同 id 重试仍合并（不重建不重执行）
    server._REQUEST_ID_CACHE.clear()
    e1, is_owner1 = server._claim("rid-x", "hash-x")
    assert is_owner1 is True and e1.state == "inflight"
    e2, is_owner2 = server._claim("rid-x", "hash-x")
    assert is_owner2 is False and e2 is e1  # 并发同 ID 合并同一 entry
    e3, is_owner3 = server._claim("rid-x", "hash-x")
    assert is_owner3 is False and e3 is e1  # owner 超时后仍合并，不重执行
    assert e1.state == "inflight"


def test_claim_hash_mismatch_rejects(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    server._claim("rid-y", "hash-1")
    with pytest.raises(server.InvalidArgumentError):
        server._claim("rid-y", "hash-2")


def test_idempotency_lru_cap(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    monkeypatch.setattr(server, "_REQUEST_ID_MAX", 3)
    for i in range(5):
        e, is_owner = server._claim(f"rid-{i}", f"hash-{i}")
        assert is_owner is True
        server._complete_entry(e, {"ok": True, "i": i})
    assert len(server._REQUEST_ID_CACHE) == 3
    assert server._REQUEST_ID_CACHE.get("rid-0") is None  # 最旧 done 被淘汰
    assert server._REQUEST_ID_CACHE.get("rid-4") is not None


def test_idempotency_lru_skips_inflight(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    monkeypatch.setattr(server, "_REQUEST_ID_MAX", 2)
    e1, _ = server._claim("a", "h-a")
    e2, _ = server._claim("b", "h-b")
    server._complete_entry(e1, {"ok": True})
    server._claim("c", "h-c")  # 超限：只淘汰 done(a)，保留 inflight b 与 c
    assert "a" not in server._REQUEST_ID_CACHE
    assert "b" in server._REQUEST_ID_CACHE
    assert "c" in server._REQUEST_ID_CACHE


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


# --- P1-1 Windows named mutex 防双启 ---


def _fake_win32_modules(already_exists: bool):
    """注入假 win32event/win32api：CreateMutex 按 already_exists 报 GetLastError。"""
    fake_ev = types.ModuleType("win32event")
    fake_ev.created_name = None

    def create(parent, initial, name):
        fake_ev.created_name = name
        return object()

    fake_ev.CreateMutex = create
    fake_api = types.ModuleType("win32api")
    fake_api.GetLastError = lambda: 183 if already_exists else 0  # ERROR_ALREADY_EXISTS
    return fake_ev, fake_api


def test_acquire_startup_lock_returns_handle_when_free(monkeypatch):
    fake_ev, fake_api = _fake_win32_modules(already_exists=False)
    monkeypatch.setitem(sys.modules, "win32event", fake_ev)
    monkeypatch.setitem(sys.modules, "win32api", fake_api)
    monkeypatch.setattr(server.sys, "platform", "win32")
    assert server._acquire_startup_lock(8890) is not None
    assert fake_ev.created_name == "Local\\offipy_server_8890"


def test_acquire_startup_lock_rejects_double_start(monkeypatch):
    fake_ev, fake_api = _fake_win32_modules(already_exists=True)
    monkeypatch.setitem(sys.modules, "win32event", fake_ev)
    monkeypatch.setitem(sys.modules, "win32api", fake_api)
    monkeypatch.setattr(server.sys, "platform", "win32")
    with pytest.raises(server.ServerStartError, match="在运行"):
        server._acquire_startup_lock(8890)


def test_acquire_startup_lock_noop_off_windows(monkeypatch):
    monkeypatch.setattr(server.sys, "platform", "linux")
    assert server._acquire_startup_lock(8890) is None  # 纯模块/WSL 不拉锁


# --- P1-3 PID 进程创建时间校验（防 PID 复用误杀） ---


def test_process_start_time_noop_off_windows(monkeypatch):
    monkeypatch.setattr(client.sys, "platform", "linux")
    assert client._process_start_time(999999) is None


def test_process_start_time_parses_powershell_iso(monkeypatch):
    fake = types.SimpleNamespace(stdout="2026-08-05T01:02:03.0000000+00:00")
    monkeypatch.setattr(client.sys, "platform", "win32")
    monkeypatch.setattr("offipy.client.subprocess.run", lambda *a, **k: fake)
    from datetime import datetime, timezone

    expected = datetime(2026, 8, 5, 1, 2, 3, tzinfo=timezone.utc).timestamp()
    assert abs(client._process_start_time(1) - expected) < 1


def test_pid_file_matches_rejects_reused_pid(monkeypatch, tmp_path):
    token = "t"
    pid = 5555
    digest = hashlib.sha256(token.encode()).hexdigest()
    (tmp_path / "server.pid").write_text(
        json.dumps({"port": client.PORT, "pid": pid, "token_sha256": digest, "started_at": 1000.0}),
        encoding="utf-8",
    )
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    monkeypatch.setattr(client, "_token", lambda: token)
    monkeypatch.setattr(client, "_process_start_time", lambda p: 9999.0)  # 偏差 > 窗口
    assert client._pid_file_matches(pid) is False  # PID 复用 → 拒认归属
    monkeypatch.setattr(client, "_process_start_time", lambda p: 1000.5)  # 吻合
    assert client._pid_file_matches(pid) is True
    monkeypatch.setattr(client, "_process_start_time", lambda p: None)  # 查不到 → 不校验
    assert client._pid_file_matches(pid) is True


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
