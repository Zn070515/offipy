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
