"""操作日志（P2-3）：JSONL append + 线程/进程双锁 + 体积轮转。

记录 server 每次 op 的 {ts, session_id, app, op, ok, error_code, duration_ms,
resource_id}——args 一律不落盘（脱敏，不写原始敏感值）。CLI `offipy log` 读取。

并发安全：写前拿模块级线程锁（进程内多线程串行）+ 文件字节锁
（Windows msvcrt / POSIX fcntl，跨进程互斥）；轮转在线程锁 + 跨进程文件锁
内完成——改名期间既无本进程写入，也不会有其它进程持有 fd，Windows 下
rename 不会因并发句柄静默失败（#48）。
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


def configure(port: int) -> None:
    """P2-2 多实例：操作日志按端口隔离（默认端口沿用 oplog.jsonl）。"""
    global _PATH
    name = "oplog.jsonl" if port == 8890 else f"oplog-{port}.jsonl"
    _PATH = user_data_dir() / name


def log_path() -> Path:
    """当前日志文件路径（测试可 monkeypatch _PATH 重定向）。"""
    return _PATH


@contextlib.contextmanager
def _file_lock(path: Path):
    """跨进程写锁：Windows msvcrt 字节锁 / POSIX fcntl。

    锁在旁路 `.lock` 文件上，不锁数据文件本身——Windows 下数据文件只要还有
    打开的 fd（包括自己）就 rename 失败（WinError 32），轮转必须在无 fd
    占用时改名；持锁方只把数据 fd 短暂打开用于写入，轮转期间其它进程被
    .lock 挡在门外，rename 才有保障（#48）。
    """
    lock_file = path.with_name(path.name + ".lock")
    with Path(lock_file).open("a+b") as f:
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
            yield
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
    with _THREAD_LOCK, _file_lock(path):
        with Path(path).open("ab") as f:
            f.write(line.encode("utf-8") + b"\n")
        _rotate(path)  # #48：在跨进程锁内轮转，改名无 fd 占用才不被打断


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

    def _parse(line: str):
        try:
            return json.loads(line)
        except ValueError:
            return None

    return [parsed for parsed in map(_parse, lines) if parsed is not None]
