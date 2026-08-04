# scripts/fetch_icons.py
"""抓取 Phosphor（fill 权重）+ Lucide 图标资产到 src/offipy/assets/icons/。

来源：
  Phosphor  https://github.com/phosphor-icons/core    MIT  （assets/fill/*.svg）
  Lucide    https://github.com/lucide-icons/lucide    ISC  （icons/*.svg）

用法：uv run python scripts/fetch_icons.py
网络：默认 opener 自动用系统代理（环境变量 https_proxy 优先，否则 Windows 注册表
      系统代理）；出站受限时先 export https_proxy=<代理地址>
"""

from __future__ import annotations

import io
import json
import re
import tarfile
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "src" / "offipy" / "assets" / "icons"

SOURCES = {
    "phosphor": {
        "repo": "phosphor-icons/core",
        "url": "https://codeload.github.com/phosphor-icons/core/tar.gz/refs/heads/main",
        "pattern": r"^[^/]+/assets/fill/([^/]+)\.svg$",
        "license": "MIT",
    },
    "lucide": {
        "repo": "lucide-icons/lucide",
        "url": "https://codeload.github.com/lucide-icons/lucide/tar.gz/refs/heads/main",
        "pattern": r"^[^/]+/icons/([^/]+)\.svg$",
        "license": "ISC",
    },
}


def _opener():
    """默认 opener：urllib 自动用系统代理（环境变量 https_proxy 优先，否则
    Windows 注册表系统代理）。

    注意：不能像 client.py 那样 ProxyHandler({}) 直连——那是针对本地回环劫持；
    这里出站抓取 GitHub 必须走代理才能出网。
    """
    return urllib.request.build_opener()


def _download(url: str) -> bytes:
    print(f"[fetch] {url}")
    try:
        with _opener().open(url, timeout=300) as resp:
            return resp.read()
    except urllib.error.HTTPError as exc:
        print(f"[error] {url} → HTTP {exc.code} {exc.reason}")
        print("  必要时 export https_proxy=<代理地址>")
        raise
    except urllib.error.URLError as exc:
        print(f"[error] {url} → {exc.reason}")
        print("  必要时 export https_proxy=<代理地址>")
        raise


def _resolve_sha(repo: str) -> str:
    """通过 GitHub API 解析 main 分支当前 commit sha（走系统代理）。"""
    url = f"https://api.github.com/repos/{repo}/commits/main"
    print(f"[resolve-sha] {url}")
    with _opener().open(url, timeout=120) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    sha = data["sha"]
    if not re.fullmatch(r"[0-9a-f]{40}", sha):
        raise ValueError(f"unexpected sha: {sha!r}")
    return sha


def _extract_commit(top: str) -> str:
    m = re.search(r"-[0-9a-f]{7,40}$", top)
    return m.group(0)[1:] if m else "unknown"


def _fetch_set(name: str, spec: dict, url: str) -> tuple[str, int]:
    dest = ASSETS / name
    dest.mkdir(parents=True, exist_ok=True)
    blob = _download(url)
    rx = re.compile(spec["pattern"])
    commit = "unknown"
    count = 0
    with tarfile.open(fileobj=io.BytesIO(blob), mode="r:gz") as tf:
        for m in tf.getmembers():
            if not m.isfile():
                continue
            hit = rx.match(m.name)
            if not hit:
                continue
            fname = hit.group(1)
            if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", fname):
                continue
            f = tf.extractfile(m)
            if f is None:
                continue
            (dest / f"{fname}.svg").write_bytes(f.read())
            count += 1
        tops = {m.name.split("/", 1)[0] for m in tf.getmembers()}
        if tops:
            commit = _extract_commit(min(tops))
    if count == 0:
        print(f"[warn] {name}: pattern 零命中，上游目录结构可能变了（pattern={spec['pattern']}）")
    # LICENSE 文本（从 tar 里复制官方许可文件）
    lic = None
    with tarfile.open(fileobj=io.BytesIO(blob), mode="r:gz") as tf:
        for m in tf.getmembers():
            if not m.isfile() or not re.search(r"(^|/)LICENSE(\.[^/]*)?$", m.name, re.I):
                continue
            f = tf.extractfile(m)
            if f is not None:
                lic = f.read().decode("utf-8", errors="replace")
                break
    (ASSETS / f"LICENSE-{name}.txt").write_text(
        lic or f"# {name} icon set\nLicense: {spec['license']}\nSource: {url}\n",
        encoding="utf-8",
    )
    print(f"  → {count} icons, commit={commit}, license={spec['license']}")
    return commit, count


def main() -> int:
    ASSETS.mkdir(parents=True, exist_ok=True)
    manifest = {}
    for name, spec in SOURCES.items():
        url = spec["url"]
        resolved = None
        try:
            resolved = _resolve_sha(spec["repo"])
            url = f"https://codeload.github.com/{spec['repo']}/tar.gz/{resolved}"
        except Exception as exc:
            print(
                f"[warn] {name}: 解析 main commit sha 失败（{exc}），"
                "回退 main 分支 URL，commit=unknown"
            )
        commit, count = _fetch_set(name, spec, url)
        manifest[name] = {
            "source": url,
            "commit": resolved or commit,
            "count": count,
            "fetched": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "license": spec["license"],
            "filter": spec["pattern"],
        }
    (ASSETS / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (ASSETS / "README.md").write_text(
        """# offipy 图标资产

从官方 repo 抓取（fetch 后覆盖本目录）：

- **phosphor/**  Phosphor fill 权重（256 viewBox，fill 模式）  MIT
- **lucide/**    Lucide 全量（24 viewBox，stroke 模式）      ISC

来源与版本见 `manifest.json`。更新：`uv run python scripts/fetch_icons.py`。

`ph:` / `lu:` 前缀对应本目录两个子目录；图标名 = 文件名（不含扩展名）。
""",
        encoding="utf-8",
    )
    print(f"manifest: {ASSETS / 'manifest.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
