"""RPC op 延迟基准（P2-7）：冷调用（首次触发 COM 启动）vs 热调用。

依赖本机 Office + 常驻 server。用法：
    uv run python scripts/bench/bench_ops.py

输出 Markdown 表（可直接并入 docs/benchmarks.md）。结束自动 quit 各 App。

CI 门禁（#107）：office-tests.yml 的真机 job 在 COM 测试后跑本脚本作为
轻档门禁，验证脚本不坏 + 指标能产出；阈值不在此设（自托管机器波动大）。
docs/benchmarks.md 是人工维护的静态快照——跑出结果后手动更新，别让快照
与实测漂移成第二个不透明数字。
"""

from time import perf_counter

from offipy import client

ROUNDS = 5
WARMUP = 2


def _time(fn) -> float:
    t0 = perf_counter()
    fn()
    return perf_counter() - t0


def bench(_app: str, _op: str, label: str, fn) -> tuple[str, float, float]:
    cold = _time(fn)  # 该 App 首个 op：COM 应用冷启动
    for _ in range(WARMUP):
        fn()
    hot = min(_time(fn) for _ in range(ROUNDS))
    return label, cold, hot


def main() -> None:
    client.ensure_server()
    rows = []
    # Excel 首 op 触发应用启动
    rows.append(
        bench("excel", "new_book", "excel.new_book", lambda: client.call("excel", "new_book"))
    )
    rows.append(
        bench(
            "excel",
            "set_cell",
            "excel.set_cell",
            lambda: client.call("excel", "set_cell", sheet=1, cell="A1", value=1),
        )
    )
    rows.append(
        bench(
            "excel",
            "read_range",
            "excel.read_range",
            lambda: client.call("excel", "read_range", sheet=1, range_addr="A1:A1"),
        )
    )
    # Word
    rows.append(bench("word", "new_doc", "word.new_doc", lambda: client.call("word", "new_doc")))
    rows.append(
        bench(
            "word",
            "write_line",
            "word.write_line",
            lambda: client.call("word", "write_line", text="hello"),
        )
    )
    rows.append(
        bench(
            "word",
            "read_doc_text",
            "word.read_doc_text",
            lambda: client.call("word", "read_doc_text"),
        )
    )
    # PPT
    rows.append(bench("ppt", "new_pres", "ppt.new_pres", lambda: client.call("ppt", "new_pres")))
    rows.append(bench("ppt", "add_slide", "ppt.add_slide", lambda: client.call("ppt", "add_slide")))
    rows.append(
        bench(
            "ppt",
            "set_title",
            "ppt.set_title",
            lambda: client.call("ppt", "set_title", slide_idx=1, text="t"),
        )
    )
    rows.append(
        bench(
            "ppt",
            "read_slide_texts",
            "ppt.read_slide_texts",
            lambda: client.call("ppt", "read_slide_texts"),
        )
    )

    client.call("excel", "quit")
    client.call("word", "quit")
    client.call("ppt", "quit")

    print("| op | 首次（冷）ms | 热调用 ms |")
    print("| --- | --- | --- |")
    for label, cold, hot in rows:
        print(f"| {label} | {cold * 1000:.1f} | {hot * 1000:.2f} |")


if __name__ == "__main__":
    main()
