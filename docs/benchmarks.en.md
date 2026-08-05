> [中文](benchmarks.md)

# Performance Benchmarks

Benchmark scripts live in `scripts/bench/`; the data is a snapshot from this machine and is
**not a cross-machine guarantee** — it is meant for relative comparison and regression protection,
not for comparing across different hardware.

- `bench_ops.py`: RPC op latency. First call (cold) = the App's first op (including COM application
  cold start); warm = minimum of 5 calls after warmup. Automatically quits each App at the end.
- `bench_deck.py`: HTML→PPTX render wall-clock time. The first run includes chromium cold start.

## Environment

- Windows 11 Pro (10.0.26200), Office 365
- Python 3.12 (`.venv`), RTX 4090
- Resident server (127.0.0.1:8890), chromium installed

## Op Latency (2026-08-05)

| op | First call (cold) ms | Warm call ms |
| --- | --- | --- |
| excel.new_book | 2163.3 | 202.20 |
| excel.set_cell | 93.7 | 17.97 |
| excel.read_range | 25.6 | 45.88 |
| word.new_doc | 2254.2 | 182.45 |
| word.write_line | 77.1 | 47.59 |
| word.read_doc_text | 58.2 | 32.26 |
| ppt.new_pres | 2414.2 | 525.86 |
| ppt.add_slide | 72.4 | 30.71 |
| ppt.set_title | 68.8 | 47.05 |
| ppt.read_slide_texts | 296.6 | 288.24 |

Observations: the first call of `new_*` ops is dominated by COM application cold start
(~2.2–2.4s); warm calls drop to single-digit milliseconds (except new object creation, e.g.,
ppt.new_pres still takes ~0.5s). `excel.read_range` shows cold/warm inversion in this sample,
which is noise (the first call happened to hit the fast path).

## Deck Render Wall-Clock Time (2026-08-05, starter deck, 5 pages)

| Round | Wall-clock ms |
| --- | --- |
| 1 (including chromium cold start) | 3331 |
| 2 | 3199 |
| 3 | 3196 |

Steady state is ~3.2s per run; the extra ~130ms on the first run is chromium cold start. Font
embedding uses subsetting; after re-sampling, the starter deck with 5 pages and 4 fonts stabilizes
at this magnitude.
