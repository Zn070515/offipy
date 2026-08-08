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
    """真实线程 server + 假 get_app/dispatch（隔离 COM），并预启动 worker。"""
    calls = []

    def fake_get_app(name):
        return object()

    def fake_dispatch(app, op, args, app_name):
        from offipy.exceptions import (
            ComOperationError,
            FileConflictError,
            InvalidArgumentError,
            TargetNotFoundError,
        )

        calls.append((app_name, op, args))
        if op == "boom":
            raise ValueError("boom op")
        if op == "raises_invalid":
            raise InvalidArgumentError("bad arg")
        if op == "raises_notfound":
            raise TargetNotFoundError("no target")
        if op == "raises_conflict":
            raise FileConflictError("file exists")
        if op == "raises_com":
            raise ComOperationError("com failed", hresult=-2147417848)
        if op == "slow":
            time.sleep(0.4)
            return "slow-done"
        if op == "order":
            return len(calls)
        return {"fake": True}

    monkeypatch.setattr(server, "get_app", fake_get_app)
    monkeypatch.setattr(server, "dispatch", fake_dispatch)
    monkeypatch.setitem(
        server._OPS,
        "ppt",
        server._OPS["ppt"]
        | {
            "boom",
            "raises_invalid",
            "raises_notfound",
            "raises_conflict",
            "raises_com",
            "slow",
            "order",
        },
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


def test_error_trace_redacted_no_path_or_source(srv):
    # D5：trace 脱敏——只留异常链 type+message，不含 File/行号/源码行（服务器信息泄露）
    status, body = _post(srv, {"app": "ppt", "op": "boom"}, token=TOKEN)
    assert status == 500
    assert body["trace"] == ["ValueError: boom op"]
    for line in body["trace"]:
        assert "File " not in line and not line.startswith("  ")


def test_safe_trace_redacts_chained_exception(monkeypatch):
    # D5 单测：异常链逐条只保留 type+message，绝不带 File/行号/源码行
    try:
        try:
            raise ValueError("inner")
        except ValueError as inner:
            raise RuntimeError("outer") from inner
    except RuntimeError as e:
        trace = server._safe_trace(e)
    assert trace == ["RuntimeError: outer", "ValueError: inner"]
    assert not any("File " in line for line in trace)
    assert not any(line.startswith("  ") for line in trace)


def test_redact_message_replaces_paths_and_doc_id():
    # #67：消息级脱敏——绝对路径（Windows/POSIX/UNC）与 doc_id 值全部替换
    assert (
        server._redact_message("open C:\\Users\\Alice\\AppData\\Roaming\\offipy\\tmp\\doc.pptx")
        == "open [REDACTED]"
    )
    assert server._redact_message("/home/xu/.cache/offipy/art.json") == "[REDACTED]"
    assert server._redact_message(r"\\server\share\doc.pptx") == "[REDACTED]"
    assert server._redact_message("target doc_id: abc-123 not found") == (
        "target [REDACTED] not found"
    )
    # 非路径/非标识原文不动：相对路径、URL、普通词
    assert server._redact_message("C:foo relative\\x") == "C:foo relative\\x"
    assert server._redact_message("see http://example.com/a") == "see http://example.com/a"


def test_error_result_and_trace_redact_message_content():
    # #67：c5ef6be 只删 traceback 帧，_error_result.error / _safe_trace 消息原文仍带
    # 绝对路径与 doc_id。现在两者都经 _redact_message 脱敏。
    msg = (
        "open C:\\Users\\Alice\\AppData\\Roaming\\offipy\\tmp\\doc.pptx "
        "/home/xu/.cache/offipy/art.json doc_id: abc-123"
    )
    try:
        try:
            raise ValueError("inner " + msg)
        except ValueError as inner:
            raise RuntimeError(msg) from inner
    except RuntimeError as e:
        res = server._error_result("word.new_doc", e)
        trace = server._safe_trace(e)
    for field in (res["error"], *trace):
        assert "C:\\Users\\Alice" not in field
        assert "/home/xu" not in field
        assert "abc-123" not in field
    assert res["error"].count("[REDACTED]") == 3
    assert all("[REDACTED]" in line for line in trace)


def test_call_error_400_with_domain_code(srv):
    # D4：invalid_argument 领域异常 → 400（非 500），error_code 仍往返
    status, body = _post(srv, {"app": "ppt", "op": "raises_invalid"}, token=TOKEN)
    assert status == 400
    assert body["ok"] is False
    assert body["error_code"] == "invalid_argument"  # 领域异常 code 往返
    assert body["resource_id"] is None


def test_call_error_status_mapping_by_code(srv):
    # D4：error_code → HTTP 状态码（client 仍从 body 映射，状态码只影响监控观感）
    cases = {
        "raises_notfound": (404, "target_not_found"),
        "raises_conflict": (409, "file_conflict"),
        "raises_com": (502, "com_operation"),
        "boom": (500, "internal"),
    }
    for op, (status, code) in cases.items():
        st, body = _post(srv, {"app": "ppt", "op": op}, token=TOKEN)
        assert st == status, f"{op}: 期望 {status}，实际 {st}"
        assert body["ok"] is False and body["error_code"] == code
    # com_operation 带 hresult 往返
    st, body = _post(srv, {"app": "ppt", "op": "raises_com"}, token=TOKEN)
    assert body["hresult"]


# --- 操作日志（P2-3）：每次 op 后落一条 ---


def test_oplog_written_on_success(srv):
    from offipy import oplog

    status, _ = _post(srv, {"app": "ppt", "op": "get_target"}, token=TOKEN)
    assert status == 200
    entries = oplog.read()
    assert any(e["app"] == "ppt" and e["op"] == "get_target" and e["ok"] is True for e in entries)
    assert all(e["session_id"] for e in entries)


def test_oplog_written_on_error(srv):
    from offipy import oplog

    status, body = _post(srv, {"app": "ppt", "op": "boom"}, token=TOKEN)
    assert status == 500
    entries = oplog.read()
    assert any(
        e["app"] == "ppt"
        and e["op"] == "boom"
        and e["ok"] is False
        and e["error_code"] == "internal"
        for e in entries
    )


def test_oplog_session_id_matches_status(srv):
    from offipy import oplog

    status, body = _get(srv, "/status", token=TOKEN)
    sid = body["result"]["session_id"]
    assert sid
    _post(srv, {"app": "ppt", "op": "order"}, token=TOKEN)
    _post(srv, {"app": "ppt", "op": "boom"}, token=TOKEN)
    assert oplog.read()  # 有记录
    assert all(e["session_id"] == sid for e in oplog.read())
