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


def test_redact_message_covers_spaced_windows_paths():
    # #75：_PATH_RE 的 Windows 模式 [^\s:;\"']* 遇空格截断——C:\Users\John Doe\...
    # 只脱 C:\Users\John，用户名尾部/临时目录/文件名全漏。路径内空格合法，
    # 应以引号/字符串结束定界，不把空格列进排除集。
    out = server._redact_message(
        r"源文件不存在: C:\Users\John Doe\AppData\Local\Temp\offipy-abc\financial.pptx"
    )
    assert out == "源文件不存在: [REDACTED]"
    for leaked in ("Doe", "AppData", "financial", "John"):
        assert leaked not in out


def test_redact_message_covers_non_whitelist_posix_roots():
    # #75：_PATH_RE 的 POSIX 模式枚举 13 个白名单根——/data /workspace /app 等
    # 数据/部署常见根原样泄漏。改「以 / 起头、≥2 段」的绝对路径形态，不再枚举根。
    out = server._redact_message("No such file: /data/offipy/docs/2026Q1.pptx")
    assert "[REDACTED]" in out
    for leaked in ("/data", "offipy", "2026Q1"):
        assert leaked not in out
    out2 = server._redact_message(
        "/workspace/exp/financial/2026Q1_report.pptx (doc_id: report-2026q1)"
    )
    assert "workspace" not in out2
    assert "report-2026q1" not in out2


def test_redact_message_covers_file_url_forms():
    # #78：0c02694 修 #75 时把 POSIX lookbehind 收紧为 (?<![:/@\w])——file:///Users 等
    # / 前是 / 的形态全被挡死，0.14.3 可脱变 0.14.4 泄漏。file:// 是桌面/前端工具链
    # 标准路径形态，须专用分支覆盖任意平台（/Users /home /data /C:）。
    cases = [
        ("file:///Users/alice/docs/x.pptx", "alice"),
        ("file:///home/bob/data/y.png", "bob"),
        ("file:///data/offipy/docs/2026Q1.pptx", "offipy"),
        ("file:///C:/Users/secret/x.pptx", "secret"),
    ]
    for msg, leaked in cases:
        out = server._redact_message(f"err: {msg}")
        assert "[REDACTED]" in out, (msg, out)
        assert leaked not in out, (msg, out)
    # 对照：http(s) URL 不误伤（file:// 分支与 POSIX lookbehind 双保险）
    assert server._redact_message("详见 https://host/path/page") == ("详见 https://host/path/page")


def test_redact_message_covers_tight_windows_paths():
    # #80：Windows 分支 \b[A-Za-z]: 词边界在盘符前紧贴 \w（中文/英文/数字/下划线）
    # 时失效——\w 与 \w 之间无边界，整条路径原样泄漏。去掉 \b + 盘符后 [\\/] 处
    # (?!\/) 排除 :// scheme：无分隔拼接形态全脱，http(s) URL 不误伤。
    cases = [
        ("找不到C:\\Users\\secret\\x.pptx", "找不到[REDACTED]"),
        ("打开C:/Users/secret/x.pptx", "打开[REDACTED]"),
        ("pathC:\\Users\\secret\\x.pptx", "path[REDACTED]"),
        ("v2C:\\Users\\secret\\x.pptx", "v2[REDACTED]"),
        ("root_C:\\Users\\secret\\x.pptx", "root_[REDACTED]"),
        ("err: C:\\Users\\secret\\x.pptx", "err: [REDACTED]"),  # 空格分隔对照仍脱
    ]
    for raw, expected in cases:
        out = server._redact_message(raw)
        assert out == expected, (raw, out)
        assert "secret" not in out and "Users" not in out
    # 对照：URL scheme 双斜杠不误伤（http:// 的 p:/、https:// 的 s:/ 均被 (?!\/) 挡下）
    assert server._redact_message("see http://example.com/a") == "see http://example.com/a"
    assert server._redact_message("详见 https://host/path/page") == "详见 https://host/path/page"


def test_redact_message_does_not_swallow_trailing_text():
    # #79：Windows 分支 [^\"']* 贪婪吞掉路径后尾随业务文本（已保存 C:\...pptx 到桌面 →
    # 已保存 [REDACTED]，尾随「到桌面」丢失）；POSIX 分支把 ../ 相对段误当绝对路径
    # （../docs/x.html → ..[REDACTED] 残根）。路径后正常文本应保留、相对路径不误脱。
    out = server._redact_message(r"已保存 C:\Users\secret\financial\2026Q1.pptx 到桌面")
    assert out == r"已保存 [REDACTED] 到桌面"
    assert "到桌面" in out
    out2 = server._redact_message(r"打开 'C:\Users\secret\x.pptx' 成功")
    assert out2 == r"打开 '[REDACTED]' 成功"
    out3 = server._redact_message("/data/offipy/docs/x.html 已转换")
    assert "已转换" in out3
    for rel in ("../docs/x.html", "../../a/b/c.html", "./style.css"):
        raw = f"err: {rel}"
        assert server._redact_message(raw) == raw, rel


def test_redact_message_windows_path_then_colon_url_time():
    # #81：Windows 路径后跟冒号/URL/端口/时间时不再漏脱（#80 lookahead 只认 CJK/引号/行尾）
    cases = [
        (
            r"C:\Users\foo\Desktop\report.pptx: 系统找不到指定的文件",
            r"[REDACTED]: 系统找不到指定的文件",
        ),
        (
            r"load failed C:\Users\foo\mod\a.py at https://example.com/x",
            r"load failed [REDACTED] at https://example.com/x",
        ),
        (r"C:\Users\foo\a.pptx:8080", r"[REDACTED]:8080"),
        (r"C:\Users\foo\a.pptx 12:30:45", r"[REDACTED] 12:30:45"),
    ]
    for raw, expected in cases:
        out = server._redact_message(raw)
        assert out == expected, (raw, out)


def test_redact_message_windows_path_keeps_english_tail():
    # #82：路径后跟英文尾巴/管道时只脱路径，保留排障上下文（与 #79 的「宁多勿漏」收敛）
    assert server._redact_message(r"Rendered C:\Users\foo\Desktop\a.pptx successfully in 1.2s") == (
        r"Rendered [REDACTED] successfully in 1.2s"
    )
    assert server._redact_message(r"C:\Users\foo\a.pptx | python x.py") == (
        r"[REDACTED] | python x.py"
    )
    # 对照 #81：路径后跟 URL 仍脱路径、保留 URL（漏脱与吞尾是一体两面）
    assert server._redact_message(r"C:\Users\foo\a.pptx http://x.com") == (
        r"[REDACTED] http://x.com"
    )


def test_redact_message_posix_does_not_hit_protocol_tags():
    # #89：POSIX 分支不误伤协议/版本标签——会话式 Remote*/CLI/HTTP 文案保持原样
    msg = (
        "未知演示文稿句柄: 'pres123'（当前会话未打开；doc_id 只在同会话内有效，"
        "本地直连 Ppt() 与会话式 Remote*/CLI/HTTP 互不相通；用本会话 list_docs 核对）"
    )
    out = server._redact_message(msg)
    assert out == msg
    # 特征对照：协议标签形态不脱，真实 POSIX 绝对路径仍脱
    assert server._redact_message("/CLI/HTTP") == "/CLI/HTTP"
    assert server._redact_message("Remote*/CLI/HTTP") == "Remote*/CLI/HTTP"
    assert server._redact_message("http://x.com/CLI/HTTP") == "http://x.com/CLI/HTTP"
    assert server._redact_message("/usr/local/bin") == "[REDACTED]"


def test_reply_503_drain_caps_recv_calls():
    # #83：恶意客户端抢到并发名额后持续灌字节时，无界 drain 会让 accept 线程卡死。
    # recv 永不返回空 = 旧 while 死循环；新实现读满 _DRAIN_CAP 即放弃排空。
    class FakeReq:
        def __init__(self):
            self.calls = 0

        def sendall(self, data):
            pass

        def setblocking(self, flag):
            pass

        def recv(self, n):
            self.calls += 1
            return b"x" * n  # 永不空 → 触发旧实现无限循环

        def close(self):
            pass

    req = FakeReq()
    server.Server._reply_503(object(), req)  # 方法不读 self
    cap_calls = server._DRAIN_CAP // 65536
    assert req.calls <= cap_calls
    assert req.calls >= 1  # 至少发过一次请求、清过一次缓冲
    # 有界性声明：上限即排空预算，与常量为同一来源（防两处漂移）
    assert server._DRAIN_CAP == 1 << 20


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
    # 脱敏粒度是「整段绝对路径替换」：Windows 路径放宽（允许空格）后会把后续空格
    # 分隔的 POSIX 路径/doc_id 一并吞并（H10 宁多勿漏），不锁死 [REDACTED] 数量。
    assert "[REDACTED]" in res["error"]
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
