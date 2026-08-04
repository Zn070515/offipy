# office-kit

Live Microsoft Office automation via COM（会话式驱动）。目标：让 Claude 能独立产出**美观、符合审美、言之有物**的 Office 产物（Word / PPT / Excel）。

> 当前状态：COM 会话管线（Excel / Word / PPT）已打通；「HTML-first 可编辑 PPTX」管线推进中（背景见 [`docs/gap_analysis.md`](docs/gap_analysis.md)）。

## 特性

- **会话式常驻 server**：跨调用保持 Office 窗口存活、文档 / 工作簿 / 演示文稿状态不丢
- **三套件原子操作**：Word / Excel / PowerPoint 增删改 + 保存 / 导出 PDF
- **断连自愈**：用户关窗或 Office 退出后自动重建会话
- **HTML-first 管线（进行中）**：Claude 写 HTML 幻灯片 → 原生可编辑 `.pptx` → 实况展示 + 视觉迭代

## 环境要求

- Windows + 已安装 Microsoft Office（Word / Excel / PowerPoint）
- Python ≥ 3.10（本仓库开发环境为 3.12）

## 安装

```bash
uv venv --python 3.12 .venv
uv pip install -e .
```

## 使用

```bash
# 首次调用自动在后台拉起常驻 server；之后所有操作打到同一进程
office excel new_book
office excel set_cell --sheet 1 --cell A1 --value 100
office excel format_cell --sheet 1 --cell A1 --bold true --size 14 --bg "#38BDF8"

office word new_doc
office word write_line --text "你好，世界"

office ppt new_pres
office ppt add_slide --layout 2
office ppt set_title --slide_idx 1 --text "标题"

office quit excel
```

## 开发

```bash
uv sync --extra dev                     # 装 dev 依赖（ruff / mypy / pytest）
uv run ruff check .                     # lint
uv run ruff format --check .            # 格式
uv run mypy src/office_kit              # 类型
uv run pytest tests -q                  # 测试（COM 集成测试无 Office 自动跳过）
```

## 结构

```
src/office_kit/
  core.py     # COM 应用生命周期与会话管理
  server.py   # 常驻会话 HTTP server（持有 COM 引用）
  cli.py      # `office` 命令入口
  excel.py / word.py / ppt.py   # 三套件原子操作
  client.py   # server 的 HTTP 客户端（HTML 管线复用）*
  deck.py     # HTML → 可编辑 PPTX 管线（render/open_live/export_slides）*
tests/        # pytest
docs/         # 差距分析与实施计划
third_party/  # vendored HTML→PPTX 转换器 *
```

\* 由「HTML-first 可编辑 PPTX 管线」计划新增，见 [`docs/superpowers/plans/`](docs/superpowers/plans/)。

## License

MIT
