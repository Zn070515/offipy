# HTML-First 可编辑 PPTX 管线 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 落地第一条「HTML-first 设计引擎」可编辑管线：Claude 写 16:9 HTML 幻灯片 → 转成**原生可编辑** `.pptx`（文字是文本框，不是图）→ 用现有 COM server 在真实 PowerPoint 里打开实况展示（页面在动）→ 逐页导出 PNG 供 Claude 视觉迭代。

**Architecture:** 复用 vendored `html-to-editable-pptx`（vector-first：文字→原生文本框、几何→原生形状、渐变/阴影/滤镜→局部快照垫底；Playwright 实测 DOM 坐标 → 原生 OOXML 组装）。新增 `src/office_kit/deck.py` 编排 `render → open_live → export_slides`；把 cli.py 里的 HTTP 客户端抽成 `src/office_kit/client.py` 供 deck.py 复用；CLI 暴露 `office deck make`。COM server 只负责「打开实况 + 逐页导出 PNG + 存 PDF」，不做 HTML 转换。

**Tech Stack:** Python 3.12（`.venv`）、pywin32、python-pptx、lxml、fonttools、playwright + Chromium、Microsoft PowerPoint（COM，用于实况展示 + 视觉审计渲染）、vendored `html-to-editable-pptx`（MIT）。

**前提**：项目根 `C:\Users\16275\Desktop\OfficeForClaude`；`.venv` 已存在且 `office_kit` editable-installed；`uv` 可用；PowerPoint 已装（此前已用）；session server（8890）正在运行。

---

## Task 0: git 初始化 + 基线提交

> **已吸收**：此 Task 已由「工程化基线」完成（git init、src-layout 迁移、.gitignore、
> pyproject.toml、tests/、README/LICENSE/py.typed、GitHub Actions CI 均已落地）。
> 执行本计划时直接从 Task 1 开始，且所有文件路径已改为 `src/office_kit/`。

**Files:**
- Create: `.gitignore`
- Create: `docs/gap_analysis.md` 已存在（不动）

- [ ] **Step 1: 初始化 git（项目当前不是 git 仓库）**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
git init
git branch -m main
```

- [ ] **Step 2: 写 .gitignore**

```gitignore
.venv/
__pycache__/
*.pyc
.office-kit.log
_audit/
*.audited.html
```

- [ ] **Step 3: 基线提交**

```bash
git add .gitignore pyproject.toml office_kit docs/gap_analysis.md
git commit -m "chore: office-kit 项目基线（COM 会话 server + 差距分析）"
```

Expected: commit 成功，`git log --oneline` 显示一条。

---

## Task 1: Vendor html-to-editable-pptx + 装依赖

**Files:**
- Create: `third_party/html-to-editable-pptx/`（git clone，去嵌套 .git 后随仓库提交）
- Create: `third_party/html-to-editable-pptx/.config.local.toml`

- [ ] **Step 1: clone 工具（shallow）**

```bash
mkdir -p /c/Users/16275/Desktop/OfficeForClaude/third_party
cd /c/Users/16275/Desktop/OfficeForClaude/third_party
git clone --depth 1 https://github.com/Hasasasa/html-to-editable-pptx.git
rm -rf html-to-editable-pptx/.git
```

Expected: `third_party/html-to-editable-pptx/convert.py` 存在。

- [ ] **Step 2: 装依赖到 .venv**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
uv pip install -r third_party/html-to-editable-pptx/requirements.txt
python -m playwright install chromium
```

Expected: pip 安装成功；`python -c "import playwright, pptx; print('ok')"` 输出 `ok`。Chromium 首次下载约 150MB。

- [ ] **Step 3: 写本机偏好配置**

Write `third_party/html-to-editable-pptx/.config.local.toml`：

```toml
[fonts]
auto_install = "yes"
[cleanup]
default = "clean"
[audit]
mode = "triage"
```

`auto_install="yes"` → convert.py 自动加 `--install-user-fonts`（内嵌字体装进用户目录，PowerPoint COM 渲染才不回退）；`audit.mode="triage"` → 主 agent 看总览图分流审查。

- [ ] **Step 4: 用工具自带的回归 fixture 冒烟**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python third_party/html-to-editable-pptx/convert.py \
  third_party/html-to-editable-pptx/tests/fixtures/regression_deck.html
```

Expected: 输出日志含 `[preflight]` `[measure]` `[assemble]`，产物 `.../regression_deck.pptx` 生成；`.../regression_deck.audited.html` 生成。首次跑会拉字体缓存（若用系统白名单字体则秒过）。

- [ ] **Step 5: 提交**

```bash
git add third_party
git commit -m "feat: vendor html-to-editable-pptx（vector-first HTML→可编辑 PPTX 转换器）"
```

---

## Task 2: 抽出 src/office_kit/client.py（复用 server 调用）

**Files:**
- Create: `src/office_kit/client.py`
- Modify: `src/office_kit/cli.py`（删掉被抽走的函数，改用 import）

> **网络代理注意**：cli.py 的 `_OPENER = build_opener(ProxyHandler({}))` 一并迁到 client.py。
> 用户 VPN 会在注册表写系统代理（ProxyServer=127.0.0.1:12334、ProxyOverride 为空），
> 会把本地 127.0.0.1 回环请求劫持给代理返回 502——**本地回环永远直连**，只有出站请求才考虑走代理。
> 同理，deck 管线里 Playwright 访问本地 HTML 时，启动参数也要显式绕过 localhost/127.0.0.1。

- [ ] **Step 1: 写失败测试**

Create `tests/test_client.py`：

```python
from office_kit.client import convert_value


def test_convert_value_bool():
    assert convert_value("true") is True
    assert convert_value("false") is False


def test_convert_value_number():
    assert convert_value("42") == 42
    assert convert_value("3.14") == 3.14


def test_convert_value_none():
    assert convert_value("none") is None


def test_convert_value_str():
    assert convert_value("hello") == "hello"
```

- [ ] **Step 2: 确认测试失败**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_client.py -q
```

Expected: FAIL，`ModuleNotFoundError: No module named 'office_kit.client'`。

- [ ] **Step 3: 创建 client.py（从 cli.py 原样搬移）**

Create `src/office_kit/client.py`：

```python
"""office-kit 客户端：常驻 server 的 HTTP 调用封装。"""
import json
import os
import subprocess
import sys
import time
import urllib.request

HOST = "127.0.0.1"
PORT = 8890
SERVER_MOD = "office_kit.server"
_URL = f"http://{HOST}:{PORT}"
# 禁用系统代理：Windows 的 urllib 默认读注册表代理设置，会把本地
# 127.0.0.1 请求劫持给代理（返回 502）。这里强制直连。
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _ping() -> bool:
    try:
        with _OPENER.open(f"{_URL}/ping", timeout=1):
            return True
    except Exception:
        return False


def ensure_server():
    if _ping():
        return
    logpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".office-kit.log")
    logfile = open(logpath, "a", encoding="utf-8")
    subprocess.Popen(
        [sys.executable, "-m", SERVER_MOD, "--port", str(PORT)],
        stdout=logfile,
        stderr=logfile,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    for _ in range(600):  # 最多等 60 秒（首次 gencache 可能较慢）
        if _ping():
            return
        time.sleep(0.1)
    raise SystemExit("无法启动 office-kit server，请查看 .office-kit.log")


def call(app: str, op: str, **args):
    ensure_server()
    data = json.dumps({"app": app, "op": op, "args": args}).encode("utf-8")
    req = urllib.request.Request(
        _URL + "/call", data=data, headers={"Content-Type": "application/json"}
    )
    with _OPENER.open(req, timeout=120) as r:
        resp = json.loads(r.read().decode("utf-8"))
    if not resp.get("ok"):
        print(f"[{app}::{op}] 失败: {resp.get('error')}", file=sys.stderr)
        if resp.get("trace"):
            for line in resp["trace"]:
                print("  " + line, file=sys.stderr)
        raise SystemExit(1)
    return resp.get("result")


def convert_value(v: str):
    if v in ("true", "True"):
        return True
    if v in ("false", "False"):
        return False
    if v in ("none", "None"):
        return None
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        pass
    return v
```

- [ ] **Step 4: 重写 cli.py 引用 client**

Replace 整个 `src/office_kit/cli.py`：

```python
"""office-kit CLI：把每个 Office 原子操作映射为一条子命令。

用法：
    office excel new_book
    office excel set_cell --sheet 1 --cell A1 --value 100
    office word new_doc
    office word write_line --text "你好"
    office ppt new_pres
    office ppt add_slide --layout 2
    office deck make --html examples/decks/starter/deck.html --out out/deck.pptx
    office quit excel

首次调用会自动在后台拉起常驻 server；之后所有操作都打到同一进程，
窗口持续可见、会话状态跨调用保持。
"""
import argparse
import json

from .client import call, convert_value, ensure_server


def _parse_kwargs(tokens):
    kwargs = {}
    it = iter(tokens)
    for tok in it:
        if tok.startswith("--"):
            kwargs[tok[2:]] = convert_value(next(it))
        else:
            kwargs[tok] = True
    return kwargs


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="office", description="office-kit CLI")
    sub = p.add_subparsers(dest="app")
    for app in ("excel", "word", "ppt"):
        sp = sub.add_parser(app)
        sp.add_argument("op")
        # REMAINDER：原样捕获 --key value 形式的任意 kwargs
        sp.add_argument("kwargs", nargs=argparse.REMAINDER)
    deck = sub.add_parser("deck")
    deck.add_argument("action", choices=["make"])
    deck.add_argument("kwargs", nargs=argparse.REMAINDER)
    q = sub.add_parser("quit")
    q.add_argument("app", choices=["excel", "word", "ppt"])
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    if args.app == "quit":
        ensure_server()
        call(args.app, "quit")
        print("quit ok")
        return
    if args.app == "deck":
        from .deck import make as deck_make
        kw = _parse_kwargs(args.kwargs)
        if args.action == "make":
            html = kw.pop("html", None)
            if not html:
                raise SystemExit("用法: office deck make --html <deck.html> [--out <x.pptx>] [--no-open] [--feedback <dir>]")
            pptx = deck_make(
                html,
                out=kw.pop("out", None),
                open_live_flag=not (kw.pop("no-open", False) or kw.pop("no_open", False)),
                feedback_dir=kw.pop("feedback", None),
            )
            print(json.dumps({"pptx": pptx}, ensure_ascii=False))
        return
    kw = _parse_kwargs(args.kwargs)
    result = call(args.app, args.op, **kw)
    if result is not None:
        print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
```

> 注：`deck` 分支里延迟 import `deck_make`（模块在 Task 4 才创建），保证本任务跑测试不报 ImportError。

- [ ] **Step 5: 测试通过**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_client.py -q
```

Expected: PASS（4 passed）。

- [ ] **Step 6: 回归 CLI 还能跑**

```bash
office ppt new_pres
```

Expected: 输出空 JSON `null` 或不报错；PowerPoint 里新建了空白演示文稿（说明 call() 链路完好）。

- [ ] **Step 7: 提交**

```bash
git add src/office_kit/cli.py src/office_kit/client.py tests/test_client.py
git commit -m "refactor: 抽出 client.py 复用 server 调用，CLI 骨架保留"
```

---

## Task 3: ppt.py 增加 export_slides（逐页导出 PNG）

**Files:**
- Modify: `src/office_kit/ppt.py`（追加 `export_slides` 方法）
- Create: `tests/test_ppt_export.py`（集成测试，需存活 PowerPoint）

- [ ] **Step 1: 写失败集成测试**

Create `tests/test_ppt_export.py`：

```python
import os
import pytest
from pathlib import Path

from office_kit import core
from office_kit.client import call

pytestmark = pytest.mark.skipif(
    not core.running("ppt"),
    reason="需要存活的 PowerPoint（server 8890 持有）",
)


def test_export_slides_pngs(tmp_path):
    call("ppt", "new_pres")
    call("ppt", "add_slide", layout=1)          # title
    call("ppt", "set_title", slide_idx=1, text="Export Test")
    call("ppt", "add_slide", layout=2)          # title+body
    call("ppt", "set_title", slide_idx=2, text="Page Two")
    out_dir = str(tmp_path / "png")
    paths = call("ppt", "export_slides", out_dir=out_dir, width=1920, height=1080)
    assert len(paths) == 2
    for p in paths:
        assert Path(p).exists()
        assert Path(p).stat().st_size > 0
```

- [ ] **Step 2: 确认失败**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_ppt_export.py -q
```

Expected: FAIL，`AttributeError: 'PptApp' object has no attribute 'export_slides'`。

- [ ] **Step 3: 在 ppt.py 追加 export_slides**

Add 到 `src/office_kit/ppt.py` 的 `save_pdf` 之后：

```python
    def export_slides(self, out_dir: str, width: int = 1920, height: int = 1080):
        """把当前演示文稿每一页导出为 PNG，供 Claude 视觉迭代。"""
        pres = self.active_pres()
        os.makedirs(out_dir, exist_ok=True)
        paths = []
        for i in range(1, pres.Slides.Count + 1):
            out = os.path.join(out_dir, f"slide_{i:02d}.png")
            pres.Slides(i).Export(out, "PNG", width, height)
            paths.append(out)
        return paths
```

> `slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)` 是 PowerPoint 内置；PNG 由 PowerPoint 真实渲染，与「页面在动」看到的一致。

- [ ] **Step 4: 重启 server（常驻进程加载旧代码）**

```bash
office quit ppt
office ppt new_pres
```

Expected: server 重建 ppt 实例后 `new_pres` 正常。

- [ ] **Step 5: 测试通过**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_ppt_export.py -q
```

Expected: PASS（2 张 PNG 生成，1920×1080）。

- [ ] **Step 6: 提交**

```bash
git add src/office_kit/ppt.py tests/test_ppt_export.py
git commit -m "feat: ppt export_slides —— 逐页导出 PNG 供视觉迭代"
```

---

## Task 4: src/office_kit/deck.py 编排模块

**Files:**
- Create: `src/office_kit/deck.py`
- Create: `tests/test_deck_render.py`

- [ ] **Step 1: 写失败测试（渲染管线）**

Create `tests/test_deck_render.py`：

```python
import pytest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"
STARTER = ROOT / "examples" / "decks" / "starter" / "deck.html"

pytestmark = pytest.mark.skipif(
    not CONVERT_PY.exists(),
    reason="third_party/html-to-editable-pptx 未 vendor",
)


def test_render_produces_pptx(tmp_path):
    from office_kit.deck import render
    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    assert out.exists()
    assert out.stat().st_size > 0
```

> `examples/decks/starter/deck.html` 由 Task 6 创建。若先跑本任务，可先用 `third_party/.../tests/fixtures/regression_deck.html` 顶替 STARTER（`no_visual_audit=True` 只跑 preflight→measure→assemble→embed→self_check，不依赖 PowerPoint）。

- [ ] **Step 2: 确认失败**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_deck_render.py -q
```

Expected: FAIL，`ModuleNotFoundError: No module named 'office_kit.deck'`。

- [ ] **Step 3: 创建 deck.py**

Create `src/office_kit/deck.py`：

```python
"""HTML → 可编辑 PPTX 落地管线。

Claude 写 16:9 HTML 幻灯片 → third_party 的 vector-first 转换器（Playwright
实测 DOM 坐标）转成原生可编辑 .pptx → 通过常驻 server 在真实 PowerPoint 里
打开实况展示（页面在动），并逐页导出 PNG 供 Claude 视觉迭代。
"""
import os
import subprocess
import sys
from pathlib import Path

from .client import call, ensure_server

ROOT = Path(__file__).resolve().parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"


def _convert_cmd(html: str, out: str | None, only_slides, no_visual_audit: bool) -> list[str]:
    cmd = [sys.executable, str(CONVERT_PY), str(html)]
    if out:
        cmd += ["--out", str(out)]
    if only_slides:
        cmd += ["--only-slides", ",".join(str(i) for i in only_slides)]
    if no_visual_audit:
        cmd += ["--no-visual-audit"]
    return cmd


def render(html: str, out: str | None = None, only_slides=None,
           no_visual_audit: bool = False, timeout: int = 600) -> str:
    """跑完整转换管线，返回产出 .pptx 的绝对路径。"""
    html = os.path.abspath(html)
    if not os.path.exists(html):
        raise FileNotFoundError(html)
    cmd = _convert_cmd(html, out, only_slides, no_visual_audit)
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError(f"convert.py 失败 (exit {r.returncode})\n{r.stdout}\n{r.stderr}")
    pptx = os.path.abspath(out) if out else str(Path(html).with_suffix(".pptx"))
    if not os.path.exists(pptx):
        raise FileNotFoundError(f"转换未产出 .pptx: {pptx}\n{r.stdout}\n{r.stderr}")
    return pptx


def open_live(pptx: str) -> None:
    """在真实 PowerPoint 里打开生成的 .pptx（实况展示）。"""
    ensure_server()
    call("ppt", "open_pres", path=os.path.abspath(pptx))


def export_slides(out_dir: str, width: int = 1920, height: int = 1080) -> list[str]:
    """把当前实况演示文稿逐页导出 PNG，供 Claude 视觉迭代。"""
    ensure_server()
    return call("ppt", "export_slides", out_dir=os.path.abspath(out_dir),
                width=width, height=height)


def make(html: str, out: str | None = None, open_live_flag: bool = True,
         feedback_dir: str | None = None) -> str:
    """render → （可选）打开实况 → （可选）导出 PNG 反馈。返回 .pptx 绝对路径。"""
    pptx = render(html, out)
    if open_live_flag:
        open_live(pptx)
    if feedback_dir:
        export_slides(feedback_dir)
    return pptx
```

- [ ] **Step 4: 测试通过**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_deck_render.py -q
```

Expected: PASS（若 STARTER 缺失，先用 regression_deck.html 顶替跑通再换）。

- [ ] **Step 5: 提交**

```bash
git add src/office_kit/deck.py tests/test_deck_render.py
git commit -m "feat: deck.py 编排 render→open_live→export_slides"
```

---

## Task 5: CLI 接通 `office deck make`

> cli.py 在 Task 2 已预留 deck 分支，这里只需验收。

- [ ] **Step 1: 生成一个临时 deck 并跑通**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
office deck make \
  --html third_party/html-to-editable-pptx/tests/fixtures/regression_deck.html \
  --out /c/Users/16275/Desktop/OfficeForClaude/out/smoke.pptx \
  --no-open
```

Expected: 输出 `{"pptx": "C:\\Users\\16275\\Desktop\\OfficeForClaude\\out\\smoke.pptx"}`；文件存在。

- [ ] **Step 2: 打开实况验证**

```bash
office ppt open_pres --path /c/Users/16275/Desktop/OfficeForClaude/out/smoke.pptx
```

Expected: 真实 PowerPoint 窗口里打开 smoke.pptx，能看到 2 页内容在动。

- [ ] **Step 3: 提交（无新文件则不提交）**

```bash
git add -A
git status
```

> 只提交 office_kit 相关改动；`out/` 进 .gitignore。

---

## Task 6: 设计令牌 starter deck（美 + 审美规范）

**Files:**
- Create: `examples/decks/starter/deck.html`

- [ ] **Step 1: 写失败断言（页数 = 5）**

Append 到 `tests/test_deck_render.py`：

```python
def test_starter_deck_slide_count(tmp_path):
    from office_kit.deck import render
    out = tmp_path / "deck.pptx"
    render(str(STARTER), out=str(out), no_visual_audit=True)
    from pptx import Presentation
    assert len(Presentation(str(out)).slides) == 5
```

> 用 `python-pptx` 直接验证可编辑性：打开 .pptx 数页数，并可在后续断言 `len(slide.shapes)` 说明不是整页图。

- [ ] **Step 2: 确认失败**

```bash
python -m pytest tests/test_deck_render.py::test_starter_deck_slide_count -q
```

Expected: FAIL，`FileNotFoundError`（deck.html 不存在）。

- [ ] **Step 3: 创建 starter deck**

Create `examples/decks/starter/deck.html`：

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>Starter Deck</title>
<style>
  :root {
    --bg: #0f172a;
    --surface: #1e293b;
    --ink: #f8fafc;
    --muted: #94a3b8;
    --accent: #38bdf8;
    --pad: 96px;
    --gap: 24px;
    --radius: 16px;
    --font: "Microsoft YaHei", "PingFang SC", Arial, sans-serif;
    --kicker: 20px;
    --title: 52px;
    --body: 24px;
    --caption: 18px;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: var(--font); }
  /* 每页必须是 1920x1080 的不透明 section；data-pptx-slide 显式标记 */
  .slide {
    width: 1920px; height: 1080px; position: relative;
    overflow: hidden; background: var(--bg); color: var(--ink);
    padding: var(--pad);
  }
  .slide.light { background: #f8fafc; color: #0f172a; }
  .kicker {
    font-size: var(--kicker); letter-spacing: 4px; color: var(--accent);
    text-transform: uppercase; margin-bottom: 24px;
  }
  .title { font-size: var(--title); font-weight: 700; line-height: 1.15; margin-bottom: 32px; }
  .subtitle { font-size: var(--body); color: var(--muted); max-width: 1200px; line-height: 1.6; }
  .cards { display: flex; gap: var(--gap); margin-top: 64px; }
  .card {
    background: var(--surface); border-radius: var(--radius); padding: 40px;
    flex: 1; border-left: 6px solid var(--accent);
  }
  .card .num { font-size: 28px; color: var(--accent); font-weight: 700; margin-bottom: 12px; }
  .card .txt { font-size: var(--body); color: var(--ink); line-height: 1.6; }
  .big { font-size: 200px; font-weight: 800; color: var(--accent); line-height: 1; }
  .rule { font-size: var(--body); color: var(--muted); margin-top: 24px; }
  .footer {
    position: absolute; bottom: 56px; left: var(--pad); right: var(--pad);
    font-size: var(--caption); color: var(--muted);
    display: flex; justify-content: space-between;
  }
  ul.bullets { margin-top: 48px; font-size: var(--body); line-height: 1.8; list-style: none; }
  ul.bullets li { margin-bottom: 20px; padding-left: 36px; position: relative; }
  ul.bullets li::before { content: ""; position: absolute; left: 0; top: 18px;
    width: 14px; height: 14px; border-radius: 4px; background: var(--accent); }
</style>
</head>
<body>

<section class="slide" data-pptx-slide>
  <div class="kicker">Quarterly Review · Q2 2026</div>
  <h1 class="title">产品增长报告<br>与下季度规划</h1>
  <p class="subtitle">聚焦三个核心指标：活跃用户、留存与收入；给出下一步行动清单。</p>
  <div class="footer"><span>OfficeForClaude · 设计引擎 starter</span><span>01 / 05</span></div>
</section>

<section class="slide" data-pptx-slide>
  <div class="kicker">Agenda</div>
  <h2 class="title">今天讲三件事</h2>
  <div class="cards">
    <div class="card"><div class="num">01</div><div class="txt">增长：活跃用户月环比 +18%</div></div>
    <div class="card"><div class="num">02</div><div class="txt">留存：次月留存提升 6 个百分点</div></div>
    <div class="card"><div class="num">03</div><div class="txt">规划：下季度三件必做</div></div>
  </div>
  <div class="footer"><span>OfficeForClaude · 设计引擎 starter</span><span>02 / 05</span></div>
</section>

<section class="slide" data-pptx-slide>
  <div class="kicker">Highlight</div>
  <div class="big">+18%</div>
  <h2 class="title">月活跃用户增速创新高</h2>
  <p class="rule">连续三个季度加速，本季度达到历史峰值。</p>
  <div class="footer"><span>OfficeForClaude · 设计引擎 starter</span><span>03 / 05</span></div>
</section>

<section class="slide" data-pptx-slide>
  <div class="kicker">Plan</div>
  <h2 class="title">下季度三件必做</h2>
  <ul class="bullets">
    <li>上线推荐位改版，承接新增量</li>
    <li>完善付费转化漏斗，目标提升 5%</li>
    <li>启动企业版 beta，验证 B 端需求</li>
  </ul>
  <div class="footer"><span>OfficeForClaude · 设计引擎 starter</span><span>04 / 05</span></div>
</section>

<section class="slide light" data-pptx-slide>
  <div class="kicker">Thank You</div>
  <h1 class="title">欢迎提问与反馈</h1>
  <p class="subtitle">完整数据包与复盘文档已归档，随时可取。</p>
  <div class="footer"><span>OfficeForClaude · 设计引擎 starter</span><span>05 / 05</span></div>
</section>

</body>
</html>
```

> 审美规范来源（research 结论四）：对比度 ≥4.5:1（slate-900 底 + slate-50 字，远超）、正文 ≥18pt（body=24px）、标题 ≥24pt（title=52px）、页边距 0.5"（pad=96px ≈ 1"）、每页要点 ≤3 条。字体走系统白名单（Microsoft YaHei / Arial），不触发 Google Fonts 下载，首跑秒过。

- [ ] **Step 4: 测试通过**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_deck_render.py -q
```

Expected: PASS（含 `test_starter_deck_slide_count`，python-pptx 打开 .pptx 数出 5 页）。

- [ ] **Step 5: 提交**

```bash
git add examples/decks/starter/deck.html tests/test_deck_render.py
git commit -m "feat: 设计令牌 starter deck —— 5 页 16:9，系统字体，审美规范可量化"
```

---

## Task 7: 端到端验收 + 视觉迭代闭环

**Files:**
- Create: `tests/test_deck_e2e.py`

- [ ] **Step 1: 写端到端测试**

Create `tests/test_deck_e2e.py`：

```python
import pytest
from pathlib import Path

from office_kit.client import _ping

ROOT = Path(__file__).resolve().parent.parent
CONVERT_PY = ROOT / "third_party" / "html-to-editable-pptx" / "convert.py"
STARTER = ROOT / "examples" / "decks" / "starter" / "deck.html"

pytestmark = pytest.mark.skipif(
    not CONVERT_PY.exists() or not _ping(),
    reason="需要 third_party + 存活 server(8890)",
)


def test_full_loop_render_open_export(tmp_path):
    from office_kit.deck import make
    png_dir = str(tmp_path / "png")
    pptx = make(str(STARTER), out=str(tmp_path / "deck.pptx"),
                feedback_dir=png_dir)
    assert Path(pptx).exists()
    pngs = sorted((tmp_path / "png").glob("slide_*.png"))
    assert len(pngs) == 5
    for p in pngs:
        assert p.stat().st_size > 0
```

- [ ] **Step 2: 跑通（PowerPoint 实况里应看到 starter deck 5 页在动）**

```bash
cd /c/Users/16275/Desktop/OfficeForClaude
python -m pytest tests/test_deck_e2e.py -q -s
```

Expected: PASS；屏幕上的 PowerPoint 打开 starter deck，可逐页看到设计。

- [ ] **Step 3: 视觉迭代闭环演练（手动）**

1. 在 PowerPoint 里翻看 5 页，挑一处要改的（如配色/文案）。
2. 修改的是 `<input>.audited.html`（转换器自动生成、随第一轮 render 出现在 `examples/decks/starter/` 下），**不改源 deck.html**。
3. 增量重渲只改的那一页：
```bash
office deck make --html examples/decks/starter/deck.audited.html \
  --out out/deck.pptx --only-slides 3
```
> deck.render 透传 `only_slides` 给 convert.py 的 `--only-slides`，measure 只重测该页、其它页复用缓存。
4. 再 `export_slides` 出新 PNG，喂回 Claude 复查 → 循环直到满意。
5. 交付前 `--cleanup`（convert.py 自带，config `cleanup.default="clean"`）。

- [ ] **Step 4: 提交**

```bash
git add tests/test_deck_e2e.py
git commit -m "test: deck 端到端 render→open_live→export_slides"
```

---

## Self-Review

**Spec 覆盖：** 差距分析 §2.4 演进路径第 2 条（HTML-first 设计引擎）+ §3 修正（可编辑管线直接上）→ Task 1 vendor、Task 4 render、Task 6 设计令牌全覆盖；§2.2 差距④（视觉反馈闭环）→ Task 3 export_slides + Task 7 迭代演练；COM server 保留「打开实况 + 导出 PNG」职责 → Task 3/4。内容工作流（outline→逐页）与能力补齐（图标/图表）不在本计划范围，属后续 plan。

**占位符扫描：** 无 TBD/TODO；每个改码步骤均含完整代码。

**类型一致性：** `deck.make(html, out, open_live_flag, feedback_dir)` 在 cli.py 与 e2e 测试中调用签名一致；`deck.render(html, out, only_slides, no_visual_audit, timeout)` 在 test_deck_render 中一致；`client.call` / `convert_value` 在 cli.py 与 client.py 一致；`ppt.export_slides(out_dir, width, height)` 在 ppt.py 与 deck.py 中一致。

**已知风险：**
- convert.py 首次 run 需 Chromium + 字体缓存（已 Task 1 装好；系统白名单字体不下载）。
- 视觉审计（5b）与我们的 COM server 同用一个 PowerPoint 实例：两者各自独立进程 attach，仅共享实例，不冲突（Task 1 Step 4 冒烟即验证）。
- `test_deck_render` 依赖 STARTER 存在（Task 6 才建）；任务按序执行则无碍，若乱序可先用 fixture 顶替。
