"""README 链接契约（契约3）：docs/ 与 examples/ 链接转绝对 GitHub URL。

PyPI 渲染无仓库上下文，相对链接会断链——docs/ 与 examples/ 一律用
`https://github.com/Zn070515/offipy/blob/main/<path>` 绝对地址，且目标
必须真实存在于仓库。仓库根文件（SECURITY/CONTRIBUTING/THIRD_PARTY_NOTICES、
语言切换器）维持相对链接白名单。
"""

import re
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parents[1]
_BASE = "https://github.com/Zn070515/offipy/blob/main/"

_README_PATHS = ["README.md", "README.en.md"]
_RELATIVE_OK = {
    "README.md",
    "README.en.md",
    "SECURITY.md",
    "SECURITY.en.md",
    "CONTRIBUTING.md",
    "CONTRIBUTING.en.md",
    "THIRD_PARTY_NOTICES.md",
}


def _link_targets(text: str) -> list[str]:
    return re.findall(r"\]\(([^)]+)\)", text)


def test_readme_docs_and_examples_links_are_absolute():
    for name in _README_PATHS:
        text = (_ROOT / name).read_text(encoding="utf-8")
        for target in _link_targets(text):
            if target.startswith(("docs/", "examples/")):
                pytest.fail(f"{name}: docs/examples 链接未转绝对 URL: {target}")


def test_readme_relative_links_are_whitelisted_and_exist():
    for name in _README_PATHS:
        text = (_ROOT / name).read_text(encoding="utf-8")
        for target in _link_targets(text):
            if target.startswith("http"):
                continue
            assert target in _RELATIVE_OK, f"{name}: 未知相对链接: {target}"
            assert (_ROOT / target).exists(), f"{name}: 相对链接目标不存在: {target}"


def test_readme_absolute_links_target_existing_files():
    for name in _README_PATHS:
        text = (_ROOT / name).read_text(encoding="utf-8")
        for target in _link_targets(text):
            if target.startswith(_BASE):
                rel = target[len(_BASE) :]
                assert (_ROOT / rel).exists(), f"{name}: 绝对链接目标不存在: {target}"
