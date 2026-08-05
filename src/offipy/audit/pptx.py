"""audit_pptx / compare_pptx 顶层编排：校验输入、惰性加载 python-pptx、聚合报告。

惰性 import 硬约束：本模块模块级不触发 python-pptx；`from pptx import Presentation`
只在 extract.py 读文件时发生。`from offipy import __version__` 必须放在函数内
（函数调用时 offipy 包已完全加载，避免 audit/__init__ 被 offipy/__init__ 引入时的
循环导入）。

异常语义：文件不存在/扩展名错 → InvalidArgumentError；PPTX ZIP/XML 损坏 →
ConversionError（由 python-pptx 打开时抛 BadZipFile/ParseError，此处包一层）；
缺 python-pptx → 原样 ImportError（提示 pip install offipy[deck]）。
"""

from __future__ import annotations

from pathlib import Path

from offipy.exceptions import InvalidArgumentError

from .compare import _sha256, build_diff
from .extract import _ShapeRecord, extract_presentation
from .models import (
    AUDIT_SCHEMA_VERSION,
    AuditConfig,
    PptxAuditReport,
    PptxDiffReport,
    SlideShapeSnapshot,
)
from .rules import run_rules

_PPTX_SUFFIX = ".pptx"


def _validate_pptx_path(path: str | Path, label: str) -> None:
    p = Path(path)
    if not p.exists():
        raise InvalidArgumentError(f"{label} PPTX 文件不存在: {path}")
    if p.suffix.lower() != _PPTX_SUFFIX:
        raise InvalidArgumentError(f"{label} 必须是 .pptx 文件: {path}")


def _snapshots(records: list[_ShapeRecord]) -> list[SlideShapeSnapshot]:
    """把 _ShapeRecord 压平成 HTML/SVG 渲染用的几何快照。"""
    return [
        SlideShapeSnapshot(
            slide_index=r.slide_index,
            shape_id=r.shape_id,
            name=r.name,
            shape_type=r.shape_type,
            role=r.role,
            left=r.left,
            top=r.top,
            width=r.width,
            height=r.height,
            z_order=r.z_order,
            text=r.text,
            is_rotated=r.is_rotated,
            geometry_unknown=r.geometry_unknown,
        )
        for r in records
    ]


def audit_pptx(path: str | Path, config: AuditConfig | None = None) -> PptxAuditReport:
    """对单个 PPTX 跑完整静态审计，产出 PptxAuditReport。"""
    from offipy import __version__

    _validate_pptx_path(path, "审计")
    cfg = config or AuditConfig()
    ext = extract_presentation(path)
    records = [r for slide in ext.slides for r in slide.shapes]
    findings, suppressed = run_rules(records, ext.slide_size, cfg)
    return PptxAuditReport(
        schema_version=AUDIT_SCHEMA_VERSION,
        offipy_version=__version__,
        path=str(path),
        source_sha256=_sha256(path),
        slide_size=ext.slide_size,
        slide_count=len(ext.slides),
        config=cfg,
        findings=findings,
        suppressed=suppressed,
        warnings=ext.warnings,
        shapes=_snapshots(records),
    )


def compare_pptx(
    baseline: str | Path,
    candidate: str | Path,
    *,
    audit_config: AuditConfig | None = None,
) -> PptxDiffReport:
    """基线与候选 PPTX 回归对比，产出 PptxDiffReport。"""
    from offipy import __version__

    _validate_pptx_path(baseline, "基线")
    _validate_pptx_path(candidate, "候选")
    base_ext = extract_presentation(baseline)
    cand_ext = extract_presentation(candidate)
    return build_diff(
        base_ext,
        cand_ext,
        baseline_path=str(baseline),
        candidate_path=str(candidate),
        offipy_version=__version__,
        audit_config=audit_config,
    )
