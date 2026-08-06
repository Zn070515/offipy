"""打包回归：converter 必须 vendored 进包内（P0-1），不依赖外部 third_party。"""

import re
from pathlib import Path

import offipy
from offipy.deck import CONVERT_PY


def test_convert_py_present():
    assert CONVERT_PY.exists()
    assert CONVERT_PY.is_file()


def test_convert_py_is_package_internal():
    # P0-1 回归：converter 在包内 _vendor/，不是源码树外的 third_party
    parts = CONVERT_PY.parts
    assert "_vendor" in parts
    assert "third_party" not in parts


def test_convert_scripts_present():
    # 转换器运行时依赖 scripts/ + lessons-learned 模板，缺一个都会在打包后炸
    scripts = CONVERT_PY.parent / "scripts" / "local_config.py"
    template = CONVERT_PY.parent / "references" / "lessons-learned.md.example"
    assert scripts.exists()
    assert template.exists()


def test_py_typed_present():
    # PEP 561 类型标注：py.typed 必须随 wheel 分发，否则编辑器拿不到类型
    py_typed = CONVERT_PY.parents[2] / "py.typed"
    assert py_typed.exists()
    assert py_typed.is_file()


def test_changelog_top_version_matches():
    # P0-5 对齐：CHANGELOG 首个具体版本号必须等于 __version__（未发布不重复 bump）
    changelog = Path(__file__).resolve().parent.parent / "CHANGELOG.md"
    _ver = r"^## \[(\d+\.\d+\.\d+(?:a\d+|b\d+|rc\d+)?)\]"
    m = re.search(_ver, changelog.read_text(encoding="utf-8"), re.M)
    assert m, "CHANGELOG 找不到版本标题"
    assert m.group(1) == offipy.__version__


# ---------------------------------------------------------------- 机械口径（文档 ↔ 代码）


_RULE_ID_CELL = re.compile(r"`([a-z._]+(?:/[a-z._]+)*)`")


def _doc_rule_ids(doc: str) -> set[str]:
    """从 audit 文档「稳定 rule_id」表提取 rule_id 集（margin 合并格按 / 展开）。"""
    m = re.search(
        r"^### (?:稳定 rule_id|Stable rule_id)$.*?(?=^#{2,} |\Z)",
        doc,
        re.S | re.M,
    )
    assert m, "找不到「稳定 rule_id / Stable rule_id」标题"
    ids: set[str] = set()
    for cell in _RULE_ID_CELL.findall(m.group(0)):
        parts = cell.split("/")
        ids.add(parts[0])
        prefix = parts[0].rsplit(".", 1)[0] + "."
        ids.update(prefix + p for p in parts[1:])
    return ids


def test_audit_doc_rule_tables_match_rule_ids():
    # 机械口径：docs/audit 规则表（zh+en）必须与代码 ALL_RULE_IDS 完全同步——
    # 缺一个或出现不存在的 rule_id 都直接红（防行为变化后文档漂移）
    from offipy.audit.models import ALL_RULE_IDS

    root = Path(__file__).resolve().parent.parent
    expected = set(ALL_RULE_IDS)
    for name in ("audit.md", "audit.en.md"):
        doc_ids = _doc_rule_ids((root / "docs" / name).read_text(encoding="utf-8"))
        assert doc_ids == expected, (
            f"docs/{name} 规则表与 ALL_RULE_IDS 不同步："
            f"缺 {sorted(expected - doc_ids)}，多 {sorted(doc_ids - expected)}"
        )


def test_readme_version_anchors_match():
    # 机械口径：README zh/en 的「当前版本 / Current version」锚点必须等于 __version__
    root = Path(__file__).resolve().parent.parent
    m_zh = re.search(
        r"当前版本.*?[:：]\s*(\d+\.\d+\.\d+)",
        (root / "README.md").read_text(encoding="utf-8"),
    )
    m_en = re.search(
        r"Current version.*?[:：]\s*(\d+\.\d+\.\d+)",
        (root / "README.en.md").read_text(encoding="utf-8"),
    )
    assert m_zh, "README.md 找不到「当前版本」锚点"
    assert m_en, "README.en.md 找不到 Current version 锚点"
    assert m_zh.group(1) == offipy.__version__
    assert m_en.group(1) == offipy.__version__
