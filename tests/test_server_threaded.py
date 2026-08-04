"""线程前端 + 单 COM worker + OperationResult（P1-1 / P1-3）。

起真实 ThreadingHTTPServer，monkeypatch get_app/dispatch 隔离真 COM——
worker 跑假 op，不拉 Office 进程。覆盖：
- /call 成功响应是 OperationResult 契约（ok/operation/resource_id/message/data + result 别名）
- 慢 op 排队时 /ping 不被阻塞（handler 直处理，不碰 worker）
- worker 串行保序（并发请求按入队顺序执行，结果 1..N 不重不漏）
- op 失败 → 500 + error_code（领域异常映射 / 普通异常降级 internal）
"""

import http.client
import json
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


def _post(port: int, body: dict, token: str | None = None):
    data = json.dumps(body).encode("utf-8")
    headers = {"Content-Type": "application/json"}
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
    """真实线程 server + 假 get_app/dispatch（隔离 COM），并预启动 worker。"""
    calls = []

    def fake_get_app(name):
        return object()

    def fake_dispatch(app, op, args, app_name):
        from offipy.exceptions import InvalidArgumentError

        calls.append((app_name, op, args))
        if op == "boom":
            raise ValueError("boom op")
        if op == "raises_invalid":
            raise InvalidArgumentError("bad arg")
        if op == "slow":
            time.sleep(0.4)
            return "slow-done"
        if op == "order":
            return len(calls)
        return {"fake": True}

    monkeypatch.setattr(server, "get_app", fake_get_app)
    monkeypatch.setattr(server, "dispatch", fake_dispatch)
    monkeypatch.setitem(
        server._OPS, "ppt", server._OPS["ppt"] | {"boom", "raises_invalid", "slow", "order"}
    )

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
    yield port
    srv.shutdown()
    srv.server_close()
    server._stop_worker()


# --- OperationResult 契约（P1-3） ---


def test_call_success_operation_result_shape(srv):
    status, body = _post(srv, {"app": "ppt", "op": "get_target"}, token=TOKEN)
    assert status == 200
    assert body["ok"] is True
    assert body["operation"] == "ppt.get_target"
    assert body["message"] == "ok"
    assert body["data"] == {"fake": True}
    assert body["result"] == body["data"]  # 兼容别名：新响应同时带 data 与 result
    assert "resource_id" in body  # 无真实目标时为 None，但字段必须存在


def test_call_success_serializes_scalar_data(srv):
    status, body = _post(srv, {"app": "ppt", "op": "order"}, token=TOKEN)
    assert status == 200
    assert body["data"] == 1  # int 透传，非字符串化


# --- 慢 op 不阻塞健康检查（P1-1） ---


def test_ping_not_blocked_by_slow_op(srv):
    result = {"status": None}

    def fire():
        result["status"] = _post(srv, {"app": "ppt", "op": "slow"}, token=TOKEN)

    t = threading.Thread(target=fire)
    t.start()
    time.sleep(0.05)  # 确保 slow op 已入队并被 worker 占用
    start = time.monotonic()
    status, body = _get(srv, "/ping")
    elapsed = time.monotonic() - start
    assert status == 200
    assert body["result"] == "pong"
    assert elapsed < 0.3  # 远小于 slow op 的 0.4s：handler 直处理，不进队列
    t.join(timeout=5)
    assert result["status"] is not None and result["status"][0] == 200


# --- worker 串行保序（P1-1） ---


def test_worker_serial_ordering(srv):
    threads = []
    results = [None] * 5

    def fire(i):
        results[i] = _post(srv, {"app": "ppt", "op": "order"}, token=TOKEN)

    for i in range(5):
        t = threading.Thread(target=fire, args=(i,))
        threads.append(t)
        t.start()
    for t in threads:
        t.join(timeout=10)
    datas = sorted(r[1]["data"] for r in results if r and r[0] == 200)
    # 串行执行：每次 op 独占 worker，返回各自的入队序号 1..5，不重不漏
    assert datas == [1, 2, 3, 4, 5]


# --- 失败响应带 error_code（P1-4） ---


def test_call_error_500_with_internal_code(srv):
    status, body = _post(srv, {"app": "ppt", "op": "boom"}, token=TOKEN)
    assert status == 500
    assert body["ok"] is False
    assert body["operation"] == "ppt.boom"
    assert body["error_code"] == "internal"  # 普通异常无 code → 降级
    assert "boom op" in body["error"]
    assert "trace" in body and body["trace"]


def test_call_error_500_with_domain_code(srv):
    status, body = _post(srv, {"app": "ppt", "op": "raises_invalid"}, token=TOKEN)
    assert status == 500
    assert body["ok"] is False
    assert body["error_code"] == "invalid_argument"  # 领域异常 code 往返
    assert body["resource_id"] is None
