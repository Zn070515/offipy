"""RPC op 延迟基准（P2-7）：冷调用（首次触发 COM 启动）vs 热调用。

依赖本机 Office + 常驻 server。用法：
    uv run python scripts/bench/bench_ops.py

输出 Markdown 表（可直接并入 docs/benchmarks.md）。结束自动 quit 各 App。

CI 门禁（#107）：office-tests.yml 的真机 job 在 COM 测试后跑本脚本作为
轻档门禁，验证脚本不坏 + 指标能产出；阈值不在此设（自托管机器波动大）。
docs/benchmarks.md 是人工维护的静态快照——跑出结果后手动更新，别让快照
与实测漂移成第二个不透明数字。

各 App 首 op（new_*）触发 COM 应用启动并返回 doc_id；后续 op 显式传
doc_id 路由，不依赖活动窗口解析（ActiveWorkbook/ActiveDocument/
ActivePresentation）——无交互桌面会话的真机 runner 上活动对象可能解析
不到，缺省会抛「需要显式 doc_id/expected_target/follow_active」。
"""

from time import perf_counter
from typing import Any

from offipy import client

ROUNDS = 5
WARMUP = 2


def _time(fn) -> tuple[float, Any]:
    t0 = perf_counter()
    result = fn()
    return perf_counter() - t0, result


def bench(_app: str, _op: str, label: str, fn) -> tuple[str, float, float, Any]:
    cold, cold_result = _time(fn)  # 该 App 首个 op：COM 应用冷启动（new_* 返回 doc_id）
    for _ in range(WARMUP):
        fn()
    hot = min(_time(fn)[0] for _ in range(ROUNDS))
    return label, cold, hot, cold_result


def main() -> None:
    client.ensure_server()
    rows = []
    # Excel 首 op 触发应用启动；new_book 冷调用返回 doc_id，set/read 显式路由
    new_book = bench(
        "excel", "new_book", "excel.new_book", lambda: client.call("excel", "new_book")
    )
    rows.append(new_book)
    excel_doc = new_book[3]
    rows.append(
        bench(
            "excel",
            "set_cell",
            "excel.set_cell",
            lambda: client.call("excel", "set_cell", doc_id=excel_doc, sheet=1, cell="A1", value=1),
        )
    )
    rows.append(
        bench(
            "excel",
            "read_range",
            "excel.read_range",
            lambda: client.call(
                "excel", "read_range", doc_id=excel_doc, sheet=1, range_addr="A1:A1"
            ),
        )
    )
    # Word
    new_doc = bench("word", "new_doc", "word.new_doc", lambda: client.call("word", "new_doc"))
    rows.append(new_doc)
    word_doc = new_doc[3]
    rows.append(
        bench(
            "word",
            "write_line",
            "word.write_line",
            lambda: client.call("word", "write_line", doc_id=word_doc, text="hello"),
        )
    )
    rows.append(
        bench(
            "word",
            "read_doc_text",
            "word.read_doc_text",
            lambda: client.call("word", "read_doc_text", doc_id=word_doc),
        )
    )
    # PPT
    new_pres = bench("ppt", "new_pres", "ppt.new_pres", lambda: client.call("ppt", "new_pres"))
    rows.append(new_pres)
    ppt_doc = new_pres[3]
    rows.append(
        bench(
            "ppt",
            "add_slide",
            "ppt.add_slide",
            lambda: client.call("ppt", "add_slide", doc_id=ppt_doc),
        )
    )
    rows.append(
        bench(
            "ppt",
            "set_title",
            "ppt.set_title",
            lambda: client.call("ppt", "set_title", doc_id=ppt_doc, slide_idx=1, text="t"),
        )
    )
    rows.append(
        bench(
            "ppt",
            "read_slide_texts",
            "ppt.read_slide_texts",
            lambda: client.call("ppt", "read_slide_texts", doc_id=ppt_doc, slide_idx=1),
        )
    )

    client.call("excel", "quit")
    client.call("word", "quit")
    client.call("ppt", "quit")

    print("| op | 首次（冷）ms | 热调用 ms |")
    print("| --- | --- | --- |")
    for label, cold, hot, _ in rows:
        print(f"| {label} | {cold * 1000:.1f} | {hot * 1000:.2f} |")


if __name__ == "__main__":
    main()
