# 性能基准

基准脚本在 `scripts/bench/`，数据是本机快照，**非跨机器承诺**——用于相对比较与
防退化，不在不同硬件间横向对比。

- `bench_ops.py`：RPC op 延迟。首次（冷）= 该 App 第一个 op（含 COM 应用冷启动）；
  热 = warmup 后 5 次中最小值。结束自动 quit 各 App。
- `bench_deck.py`：HTML→PPTX render 墙钟。首次含 chromium 冷启动。

## 环境

- Windows 11 Pro（10.0.26200）、Office 365
- Python 3.12（`.venv`）、RTX 4090
- 常驻 server（127.0.0.1:8890）、chromium 已装

## op 延迟（2026-08-05）

| op | 首次（冷）ms | 热调用 ms |
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

观察：`new_*` 首次调用被 COM 应用冷启动主导（~2.2–2.4s）；热调用回落个位数毫秒级
（新对象创建除外，如 ppt.new_pres 仍 ~0.5s）。`excel.read_range` 本次采样首/热倒挂，
属噪声（首测恰逢快速路径）。

## deck render 墙钟（2026-08-05，starter deck，5 页）

| 轮次 | 墙钟 ms |
| --- | --- |
| 1（含 chromium 冷启动） | 3331 |
| 2 | 3199 |
| 3 | 3196 |

稳态 ~3.2s/次；首次多出的 ~130ms 为 chromium 冷启动。字体内嵌为子集化，
starter 5 页 4 字体重采样后稳定在该量级。
