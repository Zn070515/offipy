"""操作日志（P2-3）：JSONL append + 线程/进程双锁 + 体积轮转。

记录 server 每次 op 的 {ts, session_id, app, op, ok, error_code, duration_ms,
resource_id}——args 一律不落盘（脱敏，不写原始敏感值）。CLI `offipy log` 读取。

并发安全：写前拿模块级线程锁（进程内多线程串行）+ 文件字节锁
（Windows msvcrt / POSIX fcntl，跨进程互斥）；轮转在同一把线程锁内完成，
避免改名期间有写入打到同一文件。
"""

import contextlib
import json
import sys
import threading
from datetime import datetime, timezone
from pathlib import Path

from .paths import user_data_dir

MAX_BYTES = 5 * 1024 * 1024  # ~5MB 轮转上限
_PATH: Path = user_data_dir() / "oplog.jsonl"
_THREAD_LOCK = threading.Lock()


def log_path() -> Path:
    """当前日志文件路径（测试可 monkeypatch _PATH 重定向）。"""
    return _PATH


@contextlib.contextmanager
def _file_lock(path: Path):
    """跨进程写锁：Windows msvcrt 字节锁 / POSIX fcntl；锁 + 写在同一个 fd。"""
    with open(path, "a+b") as f:
        try:
            if sys.platform == "win32":
                import msvcrt

                f.seek(0)
                msvcrt.locking(f.fileno(), msvcrt.LK_LOCK, 1)
            else:
                import fcntl

                fcntl.flock(f.fileno(), fcntl.LOCK_EX)
        except Exception:
            pass  # 拿不到跨进程锁仍继续写（单实例下线程锁已够用）
        try:
            yield f
        finally:
            with contextlib.suppress(Exception):
                if sys.platform == "win32":
                    import msvcrt

                    f.seek(0)
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(f.fileno(), fcntl.LOCK_UN)


def _rotate(path: Path) -> None:
    """超过上限：当前文件改名 .1 保留一份，体积有界。"""
    try:
        if not path.exists() or path.stat().st_size <= MAX_BYTES:
            return
        bak1 = path.with_suffix(".jsonl.1")
        bak1.unlink(missing_ok=True)  # Windows rename 不覆盖已存在目标，先删旧的 .1
        path.rename(bak1)
    except OSError:
        pass


def _append_locked(path: Path, line: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with _THREAD_LOCK:
        with _file_lock(path) as f:
            f.write(line.encode("utf-8") + b"\n")
            f.flush()
        _rotate(path)


def append(session_id: str, app: str, op: str, ok: bool, **kw) -> None:
    """写一条操作日志；失败静默（日志不得拖垮请求）。"""
    try:
        rec = {
            "ts": datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
            "session_id": session_id,
            "app": app,
            "op": op,
            "ok": ok,
            "error_code": kw.get("error_code"),
            "duration_ms": kw.get("duration_ms", 0),
            "resource_id": kw.get("resource_id"),
        }
        _append_locked(log_path(), json.dumps(rec, ensure_ascii=False))
    except Exception:
        pass


def read(tail: int | None = None) -> list[dict]:
    """读全部记录；tail 给正数则只取末尾 N 条。半行（写中读取）跳过。"""
    path = log_path()
    if not path.exists():
        return []
    lines = path.read_text(encoding="utf-8").splitlines()
    if tail is not None and tail > 0:
        lines = lines[-tail:]
    out = []
    for line in lines:
        try:
            out.append(json.loads(line))
        except ValueError:
            continue
    return out
