"""P2-3 操作日志：append / read / 轮转 / 并发写不串行。"""

import contextlib
import json
import threading

from offipy import oplog


def _point(monkeypatch, tmp_path):
    monkeypatch.setattr(oplog, "_PATH", tmp_path / "oplog.jsonl")


def test_append_and_read_roundtrip(monkeypatch, tmp_path):
    _point(monkeypatch, tmp_path)
    oplog.append("sess-1", "excel", "set_cell", True, duration_ms=3, resource_id="excel:book:Book1")
    oplog.append("sess-1", "ppt", "add_slide", False, error_code="com_operation")

    entries = oplog.read()
    assert len(entries) == 2
    e0, e1 = entries
    assert e0["session_id"] == "sess-1"
    assert e0["app"] == "excel" and e0["op"] == "set_cell"
    assert e0["ok"] is True
    assert e0["resource_id"] == "excel:book:Book1"
    assert e0["duration_ms"] == 3
    assert e0["error_code"] is None
    assert e1["ok"] is False
    assert e1["error_code"] == "com_operation"
    assert e1["resource_id"] is None


def test_read_tail(monkeypatch, tmp_path):
    _point(monkeypatch, tmp_path)
    for i in range(5):
        oplog.append("s", "excel", f"op{i}", True)
    assert [e["op"] for e in oplog.read(tail=2)] == ["op3", "op4"]
    assert [e["op"] for e in oplog.read(tail=99)] == [f"op{i}" for i in range(5)]
    assert len(oplog.read()) == 5


def test_read_missing_file(monkeypatch, tmp_path):
    monkeypatch.setattr(oplog, "_PATH", tmp_path / "nope.jsonl")
    assert oplog.read() == []


def test_rotation_keeps_log_bounded(monkeypatch, tmp_path):
    # MAX 远小于两条记录之和：每次追加后必触发轮转，当前文件最多 1 条
    _point(monkeypatch, tmp_path)
    monkeypatch.setattr(oplog, "MAX_BYTES", 250)
    for i in range(20):
        oplog.append("s", "excel", f"op{i}", True)

    current = tmp_path / "oplog.jsonl"
    backup = tmp_path / "oplog.jsonl.1"
    assert backup.exists()  # 至少触发过一次轮转
    # 轮转把当前文件约束在有界体积：超限即改名，当前 .jsonl 不存在(=0)或 ≤ MAX_BYTES
    current_bytes = current.stat().st_size if current.exists() else 0
    assert current_bytes <= oplog.MAX_BYTES
    for path in (current, backup):  # 保留的每条都是合法 JSON
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                json.loads(line)


def test_rotate_runs_inside_file_lock(monkeypatch, tmp_path):
    # #48：轮转必须在跨进程文件锁作用域内执行——此前在 _file_lock 块外调用，
    # Windows 下并发进程持有 fd 时 rename 静默失败（except OSError: pass）。
    _point(monkeypatch, tmp_path)
    monkeypatch.setattr(oplog, "MAX_BYTES", 1)  # 每次追加都触发轮转
    state = {"in_lock": False, "rotate_seen": []}
    real_lock = oplog._file_lock
    real_rotate = oplog._rotate

    @contextlib.contextmanager
    def tracking_lock(path):
        state["in_lock"] = True
        try:
            with real_lock(path) as f:
                yield f
        finally:
            state["in_lock"] = False

    def tracking_rotate(path):
        state["rotate_seen"].append(state["in_lock"])
        return real_rotate(path)

    monkeypatch.setattr(oplog, "_file_lock", tracking_lock)
    monkeypatch.setattr(oplog, "_rotate", tracking_rotate)
    for i in range(3):
        oplog.append("s", "excel", f"op{i}", True)
    assert state["rotate_seen"]
    assert all(state["rotate_seen"]), "轮转必须在 _file_lock 作用域内执行"


def test_concurrent_writes_no_loss(monkeypatch, tmp_path):
    _point(monkeypatch, tmp_path)
    N, PER = 20, 5

    def worker(i):
        for j in range(PER):
            oplog.append("s", "excel", f"op{i}-{j}", True)

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(N)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=10)

    entries = oplog.read()
    assert len(entries) == N * PER  # 并发写不丢不重
    for e in entries:
        json.dumps(e)  # 每条都是合法 JSON 行
