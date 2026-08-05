"""固定验收集（tests/fixtures/audit/）驱动：合成资产产出预期发现。

fixtures 是一次性 git 入库资产（generate_audit_fixtures.py 可再生成）：
- synthetic.pptx    一条规则一个场景 → 逐规则断言发现的严重度/豁免原因
- edge_cases.pptx   误报控制固定验收集 → 零 finding/suppressed/warning
- baseline/candidate.pptx  基线回归 → 新增越界 + 已解决贴边 + 形状变化
- deck_generated.pptx      由 offipy deck 渲染（需 chromium），只做轻量断言

CI 内文件必存在；缺失时对应测试跳过（提示先跑生成脚本）。
"""

from pathlib import Path

import pytest

from offipy.audit import (
    RULE_AUTOFIT_GROW,
    RULE_AUTOFIT_SHRINK,
    RULE_BOUNDS_OFF_CANVAS,
    RULE_BOUNDS_PARTIAL,
    RULE_MARGIN_RIGHT,
    RULE_OVERLAP_PARTIAL,
    RULE_TEXT_FIT_HORIZONTAL,
    RULE_TEXT_FIT_VERTICAL,
    Severity,
    audit_pptx,
    compare_pptx,
)

_FIX = Path(__file__).parent / "fixtures" / "audit"

_REQUIRED = ("synthetic", "edge_cases", "baseline", "candidate")
pytestmark = pytest.mark.skipif(
    not all((_FIX / f"{n}.pptx").exists() for n in _REQUIRED),
    reason="审计固定资产缺失；先跑: python tests/fixtures/audit/generate_audit_fixtures.py",
)


def _by_rule(findings) -> dict:
    return {f.rule_id: f for f in findings}


# ---------------------------------------------------------------- synthetic


def test_synthetic_intended_findings():
    report = audit_pptx(_FIX / "synthetic.pptx")
    assert report.slide_count == 1
    assert report.slide_size == pytest.approx((10.0, 7.5))

    by_id = _by_rule(report.findings)
    assert by_id[RULE_BOUNDS_PARTIAL].severity == Severity.HIGH
    assert by_id[RULE_BOUNDS_OFF_CANVAS].severity == Severity.LOW
    assert by_id[RULE_MARGIN_RIGHT].severity == Severity.LOW
    assert by_id[RULE_OVERLAP_PARTIAL].severity == Severity.MID
    assert by_id[RULE_TEXT_FIT_HORIZONTAL].severity == Severity.LOW
    assert by_id[RULE_TEXT_FIT_VERTICAL].severity == Severity.LOW
    assert by_id[RULE_AUTOFIT_GROW].severity == Severity.MID
    assert by_id[RULE_AUTOFIT_SHRINK].severity == Severity.HIGH
    assert report.max_severity == Severity.HIGH

    # 自动角色豁免全部进 suppressed 带 reason，不静默丢弃
    reasons = {s.reason for s in report.suppressed}
    assert {"full_bleed", "page_number", "intentional_containment"} <= reasons
    assert report.warnings == []


# ---------------------------------------------------------------- edge_cases（误报控制）


def test_edge_cases_zero_spurious_findings():
    """connector/hidden/rotate/flip/group 不产生任何 finding/suppressed/warning。"""
    report = audit_pptx(_FIX / "edge_cases.pptx")
    assert report.findings == []
    assert report.suppressed == []
    assert report.warnings == []
    assert report.max_severity is None


def test_edge_cases_exact_group_geometry():
    """Affine2D 组几何精确：缩放 group 子局部坐标 → 幻灯片绝对坐标。"""
    report = audit_pptx(_FIX / "edge_cases.pptx")
    by_id = {s.shape_id: s for s in report.shapes}

    g = by_id[2]  # Group 1
    gs = by_id[3]  # TextBox 2 "gs"
    g2 = by_id[4]  # Group 3（嵌套）
    deep = by_id[5]  # TextBox 4 "deep"
    assert (g.left, g.top, g.width, g.height) == pytest.approx((1, 1, 3, 2))
    # 子局部 (0,0,0.75,0.5) × scale 2 → 绝对 (1,1,1.5,1)
    assert (gs.left, gs.top, gs.width, gs.height) == pytest.approx((1, 1, 1.5, 1))
    # 嵌套组局部 (0.75,0.5,0.75,0.5) × scale 2 → 绝对 (2.5,2,1.5,1)
    assert (g2.left, g2.top, g2.width, g2.height) == pytest.approx((2.5, 2, 1.5, 1))
    # 嵌套组内子 → 与 g2 重合（局部 (0,0,0.5,0.5) 经组内 scale 3×2 撑满）
    assert (deep.left, deep.top, deep.width, deep.height) == pytest.approx((2.5, 2, 1.5, 1))

    # flipH group：子 "fl" 绕组中心镜像 → (7.5,2,0.5,1)
    fl = by_id[8]
    assert (fl.left, fl.top, fl.width, fl.height) == pytest.approx((7.5, 2, 0.5, 1))

    # 旋转 45° shape → AABB 近似标记
    rot = by_id[6]
    assert rot.is_rotated is True


# ---------------------------------------------------------------- baseline / candidate 回归


def test_compare_finds_added_and_resolved():
    diff = compare_pptx(_FIX / "baseline.pptx", _FIX / "candidate.pptx")

    added = _by_rule(diff.added_findings)
    assert added[RULE_BOUNDS_PARTIAL].severity == Severity.MID  # 新增越界
    resolved = _by_rule(diff.resolved_findings)
    assert RULE_MARGIN_RIGHT in resolved  # 已解决贴边

    assert len(diff.added_shapes) == 2
    assert len(diff.moved_shapes) == 2
    assert len(diff.resized_shapes) == 1
    assert len(diff.text_changes) == 1

    assert diff.gate_severity() == Severity.MID


# ---------------------------------------------------------------- deck_generated（需 chromium）


_DECK_REASON = (
    "deck_generated.pptx 缺失（由 offipy deck 渲染，需 chromium）；"
    "见 tests/fixtures/audit/README.md 再生成说明"
)


@pytest.mark.skipif(not (_FIX / "deck_generated.pptx").exists(), reason=_DECK_REASON)
def test_deck_generated_audits_clean():
    """deck 产物可审计、无解析 warning（真实产物，不做过度强断言）。

    已知误报记录（第 3 页，Tag 前打开页面人工复核）：
    - covered_text MID：#3 "+18%" 是 12×2.222in 的高 textbox，覆盖 #4/#5 两行
      说明文字的框——三者均为透明 TEXT_BOX（fill=BACKGROUND），视觉上各行独立
      可见，属 deck 转换器把「大数字 + 两行说明」渲染成高 textbox 的正常排版。
    - partial MID：同上 #3 与 #5 的框内叠放。
    不改进 Pair 分类：透明与否需 shape fill 信息（extract 未提取），一刀切降级
    会漏报真实遮挡；保持现状并在此固定预期，供人工验收。
    """
    report = audit_pptx(_FIX / "deck_generated.pptx")
    assert report.slide_count >= 1
    assert report.warnings == []
    assert report.shapes
    mids = [f for f in report.findings if f.severity == Severity.MID]
    assert mids  # 两条已知误报必须存在，防止被静默改掉后无人复核
    assert all(
        f.rule_id in ("geometry.overlap.covered_text", "geometry.overlap.partial") for f in mids
    )
