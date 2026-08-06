from pathlib import Path

import pytest

from offipy.art.analyze import analyze_scene
from offipy.art.models import ArtScene

FIXTURES = Path(__file__).parent / "fixtures" / "art" / "scenes"

# (fixture, profile, present_rules, forbidden_rules)
_ACCEPTANCE = [
    # 左文右图：构图平衡、有焦点、图片不失真、深灰白底高对比（不误报 low_contrast）
    # 健康场景 present 为空 []，靠测试体里「composition 维度须 assessed」断言证明规则真运行了
    # 已知：corner_cluster（大图占右上 60% 面积）与 no_accent（灰阶）会触发，均 experimental 且
    # conf 封顶 0.3、不驱动降级 —— 属 Task 20 标定的噪声项，本测试不断言也不屏蔽
    (
        "left_text_right_image.json",
        "balanced",
        [],
        [
            "art.composition.off_balance",
            "art.media.distorted_image",
            "art.hierarchy.no_focus",
            "art.color.low_contrast",
        ],
    ),
    # 黑白极简：无强调色触发 no_accent；对比度健康
    (
        "minimal_bw.json",
        "balanced",
        ["art.color.no_accent"],
        ["art.color.low_contrast", "art.typography.many_families"],
    ),
    # 高饱和封面：accent_flood 触发（experimental）；装饰条不计入
    (
        "cover_high_saturation.json",
        "event",
        ["art.color.accent_flood"],
        ["art.composition.off_balance"],
    ),
    # 通版背景：背景元素不计入 mass/margin；无失真。media 维度因加了比例正确的大图而
    # genuinely assessed（distorted_image 断言非空心）。已知噪声（Task 20 标定）：
    # corner_cluster（左侧文字区占 TL 面积多）与 no_focus（focus 中位数 quirk，见 features.py）
    (
        "full_bleed_background.json",
        "balanced",
        [],
        ["art.composition.off_balance", "art.media.distorted_image"],
    ),
]


@pytest.mark.parametrize("name,profile,present,forbidden", _ACCEPTANCE)
def test_acceptance(name, profile, present, forbidden):
    path = FIXTURES / name
    data = path.read_text(encoding="utf-8")
    report = analyze_scene(_load(data), profile=profile)
    found = {f.rule_id for s in report.slides for d in s.dimensions for f in d.findings} | {
        f.rule_id for f in report.deck_findings
    }
    for rid in present:
        assert rid in found, f"{name}: 期望出现 {rid}"
    for rid in forbidden:
        assert rid not in found, f"{name}: 期望不出现 {rid}"
    # 健康场景不能只断言「没有 Finding」——须证明规则确实运行、维度确实被评估，
    # 否则规则完全没运行也能满足 forbidden 列表（假绿）。
    if name == "left_text_right_image.json":
        comp = report.slides[0].by_dimension("composition")
        assert comp is not None and comp.status == "assessed", (
            f"{name}: composition 维度应被评估而非未运行"
        )


def _load(text):
    import json

    return ArtScene.from_dict(json.loads(text))


def test_import_offipy_does_not_load_pptx():
    """契约：import offipy 不得在 import 阶段加载 python-pptx（art/audit 纯 stdlib）。"""
    import os
    import subprocess
    import sys

    code = (
        "import sys\n"
        "import offipy.art\n"
        "import offipy.audit\n"
        "assert 'pptx' not in sys.modules, 'python-pptx 被提前加载'\n"
    )
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        env=env,
    )
    assert proc.returncode == 0, f"no-pptx 守卫失败: {proc.stderr}"
