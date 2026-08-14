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
        except OSError:  # noqa: PERF203 — 启动轮询必须每轮捕获连接失败
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
    s1, _r1 = _post(port, {**base, "args": {"a": 1}}, token=TOKEN)
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


def test_request_id_timeout_then_retry_re_executes(srv, monkeypatch):
    # D3：owner 超时(504)时释放 inflight entry——同 id 重试重建为新的 owner
    # 重新执行。不释放 → worker 卡死时该 request_id 永久 inflight，重试永远
    # merge-wait → 504 死循环。代价：慢 op 超时后重试可能重复执行一次。
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

    s1, _r1 = _post(port, body, token=TOKEN)
    assert s1 == 504  # owner 在 _CALL_TIMEOUT 内没等到 worker → 超时并释放 entry
    assert entered.wait(2)  # worker 已进入（op 执行了一次）

    gate.set()  # 放行 worker
    s2, r2 = _post(port, body, token=TOKEN)
    assert s2 == 200 and r2["data"] == "slow-done"
    assert len(calls) == 2  # 释放后重试重建 owner → 重新执行（不再永久合并）


def test_claim_inflight_merges_fresh_but_reclaims_stale(monkeypatch):
    # D3 状态机：新鲜 inflight 合并；owner 失联（started_at 超 _CALL_TIMEOUT
    # 未释放）→ 陈旧 inflight 被接管重建为新的 owner 重新执行，不永久合并。
    server._REQUEST_ID_CACHE.clear()
    e1, is_owner1 = server._claim("rid-x", "hash-x")
    assert is_owner1 is True and e1.state == "inflight"
    e2, is_owner2 = server._claim("rid-x", "hash-x")
    assert is_owner2 is False and e2 is e1  # 新鲜 inflight：并发同 ID 合并同一 entry
    # 伪造 owner 失联：把 entry 创建时刻拨回超时之前
    e1.started_at = time.monotonic() - (server._CALL_TIMEOUT + 1)
    e3, is_owner3 = server._claim("rid-x", "hash-x")
    assert is_owner3 is True and e3 is not e1  # 陈旧 inflight：接管重建为新的 owner
    assert e1.state == "inflight"
    assert server._REQUEST_ID_CACHE.get("rid-x") is e3


def test_release_inflight_only_pops_same_entry_inflight(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    e1, _ = server._claim("rid-r", "hash-r")
    e2, _ = server._claim("rid-s", "hash-s")
    server._release_inflight("rid-r", e2)  # 缓存项不是该 entry → no-op
    assert "rid-r" in server._REQUEST_ID_CACHE
    server._release_inflight("rid-r", e1)  # 匹配且仍 inflight → 释放
    assert "rid-r" not in server._REQUEST_ID_CACHE
    e3, _ = server._claim("rid-t", "hash-t")
    server._complete_entry(e3, {"ok": True})
    server._release_inflight("rid-t", e3)  # 已 done → 不释放（结果仍可回放）
    assert "rid-t" in server._REQUEST_ID_CACHE


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
    _e2, _ = server._claim("b", "h-b")
    server._complete_entry(e1, {"ok": True})
    server._claim("c", "h-c")  # 超限：只淘汰 done(a)，保留 inflight b 与 c
    assert "a" not in server._REQUEST_ID_CACHE
    assert "b" in server._REQUEST_ID_CACHE
    assert "c" in server._REQUEST_ID_CACHE


def test_idempotency_lru_move_to_end_on_hit(monkeypatch):
    # P1-5 真 LRU：done 命中 move_to_end 刷新新鲜度，淘汰从最旧开始
    server._REQUEST_ID_CACHE.clear()
    monkeypatch.setattr(server, "_REQUEST_ID_MAX", 2)
    e_a, _ = server._claim("a", "h-a")
    server._complete_entry(e_a, {"ok": True})
    e_b, _ = server._claim("b", "h-b")
    server._complete_entry(e_b, {"ok": True})
    server._claim("a", "h-a")  # done 命中 → move_to_end（a 变最新）
    server._claim("c", "h-c")  # 超限：淘汰最旧 done = b
    assert "b" not in server._REQUEST_ID_CACHE
    assert "a" in server._REQUEST_ID_CACHE
    assert "c" in server._REQUEST_ID_CACHE


def test_claim_rejects_when_inflight_cap_reached(monkeypatch):
    # P1-5 inflight 硬限：全是 inflight 时 _evict_lru 无从淘汰，必须 503
    server._REQUEST_ID_CACHE.clear()
    monkeypatch.setattr(server, "_REQUEST_MAX_INFLIGHT", 2)
    server._claim("i1", "h1")
    server._claim("i2", "h2")
    with pytest.raises(server._InflightFullError):
        server._claim("i3", "h3")


def test_idempotency_stats_counts(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    e1, _ = server._claim("s1", "h1")
    server._complete_entry(e1, {"ok": True})
    server._claim("s2", "h2")  # 留 inflight
    stats = server._idempotency_stats()
    assert stats["inflight"] == 1
    assert stats["done"] == 1
    assert stats["max"] == server._REQUEST_ID_MAX
    assert stats["max_inflight"] == server._REQUEST_MAX_INFLIGHT


def test_status_exposes_idempotency_counts(monkeypatch):
    server._REQUEST_ID_CACHE.clear()
    e1, _ = server._claim("st1", "h1")
    server._complete_entry(e1, {"ok": True})
    server._TOKEN = TOKEN
    srv = server.Server(("127.0.0.1", 0), server.Handler)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    port = srv.server_address[1]
    try:
        _wait_ready(port)
        status, body = _get(port, "/status", token=TOKEN)
        assert status == 200
        idem = body["result"]["idempotency"]
        assert idem["done"] == 1
        assert idem["inflight"] == 0
        assert idem["max_inflight"] == server._REQUEST_MAX_INFLIGHT
    finally:
        srv.shutdown()
        srv.server_close()


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


class _BlockingFullQueue:
    """put_nowait 先阻塞到 release 再抛 queue.Full——复现「owner 卡入队期间，
    同 id 非 owner 合并等待」的竞态。"""

    def __init__(self):
        self._release = threading.Event()

    def release(self):
        self._release.set()

    def put_nowait(self, item):
        if not self._release.wait(10):
            raise AssertionError("release 未在 10s 内触发")
        raise queue.Full


def test_idempotent_queue_full_rollback_wakes_non_owner(monkeypatch):
    """#45：owner 入队遇 Full 回滚时必须 set event——否则合并等待的同 id
    非 owner 线程空等 _CALL_TIMEOUT 后误报 504。回滚后应立即被唤醒、收到 busy。"""
    monkeypatch.setattr(server, "_ensure_worker", lambda: None)
    monkeypatch.setitem(server._OPS, "ppt", server._OPS["ppt"] | {"order"})
    q = _BlockingFullQueue()
    monkeypatch.setattr(server, "_COM_QUEUE", q)
    server._TOKEN = TOKEN
    srv = server.Server(("127.0.0.1", 0), server.Handler)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    port = srv.server_address[1]
    body = {"app": "ppt", "op": "order", "request_id": "r1"}
    results = {}

    def _post_a():
        results["a"] = _post(port, body, token=TOKEN)

    def _post_b():
        results["b"] = _post(port, body, token=TOKEN)

    try:
        _wait_ready(port)
        a = threading.Thread(target=_post_a, daemon=True)
        a.start()
        time.sleep(0.2)  # A 已 claim 为 owner 并卡在 put_nowait
        b = threading.Thread(target=_post_b, daemon=True)
        b.start()
        time.sleep(0.2)  # B 已 claim 并合并等待 owner 完成
        q.release()  # A 的 put_nowait 抛 Full → A 回滚并 set event → B 被唤醒
        a.join(10)
        b.join(10)
        assert "a" in results and "b" in results, "非 owner 线程必须被回滚唤醒，不能空等"
        _status_b, body_b = results["b"]
        assert body_b["error_code"] == "busy"
        assert "忙" in body_b["error"]
        assert "超时" not in body_b.get("error", "")  # 不是误导性的 504 超时
    finally:
        q.release()
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
    monkeypatch.setattr(client, "_token", lambda p: token)
    assert client._pid_file_matches(pid) is True
    assert client._pid_file_matches(pid + 1) is False  # pid 不符
    monkeypatch.setattr(client, "_token", lambda p: "other-token")
    assert client._pid_file_matches(pid) is False  # token 不符 → 拒绝


def test_client_pid_file_matches_old_plain(monkeypatch, tmp_path):
    # P0-2：旧格式纯数字无法证明 token 归属 → 一律拒绝（绝不强杀）
    (tmp_path / "server.pid").write_text("12345", encoding="utf-8")
    monkeypatch.setattr(client, "user_data_dir", lambda: tmp_path)
    monkeypatch.setattr(client, "_token", lambda p: "some-token")
    assert client._pid_file_matches(12345) is False
    assert client._pid_file_matches(1) is False


def test_remove_pid_file_if_owned(monkeypatch, tmp_path):
    monkeypatch.setattr(server, "user_data_dir", lambda: tmp_path)
    pid_file = tmp_path / "server.pid"
    # 属于本进程（pid+port 双匹配）→ 删除
    pid_file.write_text(
        json.dumps({"port": server.DEFAULT_PORT, "pid": os.getpid()}), encoding="utf-8"
    )
    server._remove_pid_file_if_owned(server.DEFAULT_PORT)
    assert not pid_file.exists()
    # 他人进程 pid → 保留（绝不误删）
    pid_file.write_text(json.dumps({"port": server.DEFAULT_PORT, "pid": 123456}), encoding="utf-8")
    server._remove_pid_file_if_owned(server.DEFAULT_PORT)
    assert pid_file.exists()
    # port 不匹配 → 保留
    pid_file.write_text(json.dumps({"port": 9999, "pid": os.getpid()}), encoding="utf-8")
    server._remove_pid_file_if_owned(server.DEFAULT_PORT)
    assert pid_file.exists()
    # 坏 JSON → 保留（无法证明归属就不删）
    pid_file.write_text("not-json", encoding="utf-8")
    server._remove_pid_file_if_owned(server.DEFAULT_PORT)
    assert pid_file.exists()


def test_close_mutex_noop_off_windows(monkeypatch):
    closed = []
    handle = types.SimpleNamespace(Close=lambda: closed.append(1))
    monkeypatch.setattr(server.sys, "platform", "linux")
    server._close_mutex(handle)
    assert closed == []  # 非 Windows 不 Close
    server._close_mutex(None)  # None 为 no-op


def test_serve_holds_mutex_until_exit(monkeypatch):
    # P0-1：mutex 句柄存局部变量持有整个 serve 生命周期；serve_forever 运行
    # 期间不释放（否则另一进程可重复启动同端口），退出 finally 才 _close_mutex。
    # 用 spy 断言 finally 编排，跨平台成立——_close_mutex 本身在非 Windows 为
    # no-op，由 test_close_mutex_noop_off_windows 单独覆盖。
    mutex = object()
    events = []

    monkeypatch.setattr(server, "_acquire_startup_lock", lambda port: mutex)
    monkeypatch.setattr(server, "_close_mutex", lambda handle: events.append("close"))
    monkeypatch.setattr(server, "oplog", types.SimpleNamespace(configure=lambda port: None))
    monkeypatch.setattr(server, "_load_token", lambda port: "t")
    monkeypatch.setattr(server, "_write_pid_file", lambda port, token: None)
    monkeypatch.setattr(server, "_ensure_worker", lambda: None)
    monkeypatch.setattr(server, "_stop_worker", lambda: None)
    monkeypatch.setattr(server, "_remove_pid_file_if_owned", lambda port: None)
    monkeypatch.setattr(server, "_TOKEN", "t")

    class FakeHTTPServer:
        def __init__(self, addr, handler):
            self.addr = addr

        def serve_forever(self):
            events.append("serve")
            raise RuntimeError("serve 退出")

        def server_close(self):
            pass

    monkeypatch.setattr(server, "Server", FakeHTTPServer)
    with pytest.raises(RuntimeError, match="serve 退出"):
        server.serve()
    assert events == ["serve", "close"]  # 运行期间未释放，退出 finally 才释放


def test_serve_bind_failure_no_pid_and_mutex_closed(monkeypatch):
    # 绑定端口失败：不写 pid 文件（不留假活），finally 仍释放 mutex（不泄漏句柄）。
    # 同 test_serve_holds_mutex_until_exit：用 spy 断言 finally 编排，跨平台成立。
    mutex = object()
    events = []
    pid_writes = []
    removes = []
    monkeypatch.setattr(server, "_acquire_startup_lock", lambda port: mutex)
    monkeypatch.setattr(server, "_close_mutex", lambda handle: events.append("close"))
    monkeypatch.setattr(server, "oplog", types.SimpleNamespace(configure=lambda port: None))
    monkeypatch.setattr(server, "_load_token", lambda port: "t")
    monkeypatch.setattr(
        server, "_write_pid_file", lambda port, token: pid_writes.append((port, token))
    )
    monkeypatch.setattr(server, "_ensure_worker", lambda: None)
    monkeypatch.setattr(server, "_stop_worker", lambda: None)
    monkeypatch.setattr(server, "_remove_pid_file_if_owned", lambda port: removes.append(port))
    monkeypatch.setattr(server, "_TOKEN", "t")

    class BindingServer:
        def __init__(self, addr, handler):
            raise OSError("端口被占用")

    monkeypatch.setattr(server, "Server", BindingServer)
    with pytest.raises(OSError, match="端口被占用"):
        server.serve()
    assert pid_writes == []  # 绑定失败未写 pid 文件（不留假活）
    assert removes == [server.DEFAULT_PORT]  # finally 仍尝试清理（归属校验兜底）
    assert events == ["close"]  # finally 释放 mutex（不泄漏句柄）


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
    monkeypatch.setattr(client, "_token", lambda p: token)
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

    def spy(path, mode, **kwargs):
        calls.append((str(path), mode))
        real(path, mode)

    monkeypatch.setattr(os, "chmod", spy)
    server._load_token(server.DEFAULT_PORT)
    assert (str(tmp_path / "token"), 0o600) in calls
