"""deck render 墙钟基准（P2-7）：首次（含 chromium 冷启动）vs 稳态。

依赖 chromium + offipy[deck]。用法：
    uv run python scripts/bench/bench_deck.py [--rounds N]

输出 Markdown 表（可直接并入 docs/benchmarks.md）。
"""

import argparse
from pathlib import Path
from time import perf_counter

ROOT = Path(__file__).resolve().parent.parent.parent
STARTER = ROOT / "examples" / "decks" / "starter" / "deck.html"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()

    from offipy.deck import render

    times = []
    for i in range(args.rounds):
        out = ROOT / "logs" / f"bench_deck_{i}.pptx"
        t0 = perf_counter()
        render(str(STARTER), out=str(out), no_visual_audit=True)
        times.append(perf_counter() - t0)
        out.unlink(missing_ok=True)

    print(f"| 轮次 | 墙钟 ms |")
    print(f"| --- | --- |")
    for i, t in enumerate(times, 1):
        tag = "（含 chromium 冷启动）" if i == 1 else ""
        print(f"| {i}{tag} | {t * 1000:.0f} |")


if __name__ == "__main__":
    main()
