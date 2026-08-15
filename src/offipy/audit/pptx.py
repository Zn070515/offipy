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
from typing import TYPE_CHECKING, cast

from offipy.exceptions import InvalidArgumentError

if TYPE_CHECKING:
    from offipy.art.models import ArtElement

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


def _to_art_elements(
    records: list[_ShapeRecord], slide_size: tuple[float, float]
) -> list[ArtElement]:
    """_ShapeRecord → 携带完整视觉证据的 ArtElement（PPTX-only 富集，#128）。

    决策 C：与 _snapshots（审计报告文本统计）职责分离，不 overload。
    坐标已 absolutize（英寸），按页宽高归一化；geometry_unknown / 缺几何跳过。
    惰性 import offipy.art（避免 audit→art 模块级环依赖）。
    """
    from offipy.art.adapters import _KIND_MAP, _infer_element_role
    from offipy.art.models import ArtColor, ArtElement, ArtTextRun

    w_in, h_in = slide_size
    width = w_in * 72.0
    height = h_in * 72.0
    out: list[ArtElement] = []
    for rec in records:
        if rec.geometry_unknown or None in (rec.left, rec.top, rec.width, rec.height):
            continue
        kind = _KIND_MAP.get(rec.shape_type.lower(), "shape")
        if rec.shape_type.lower() in ("picture", "photo"):
            kind = "image"
        runs: list[ArtTextRun] = []
        opacity: float | None = None
        for para in rec.paragraphs:
            for tr in para.runs:
                color = ArtColor(*tr.color) if tr.color is not None else None
                runs.append(
                    ArtTextRun(
                        text=tr.text,
                        font_size=tr.font_size,
                        font_size_unit="pt",
                        font_family=tr.font_name,
                        color=color,
                    )
                )
                if color is not None and color.a < 1.0:
                    opacity = min(opacity, color.a) if opacity is not None else color.a
        if rec.fill_color is not None and rec.fill_color[3] < 1.0:
            opacity = rec.fill_color[3]
        font_size = runs[0].font_size if runs else None
        background = ArtColor(*rec.fill_color) if rec.fill_color is not None else None
        role = (
            rec.role
            if rec.role != "unknown"
            else _infer_element_role(rec.name, rec.shape_type, kind)
        )
        out.append(
            ArtElement(
                element_id=f"pptx-{rec.slide_index}-{rec.shape_id}",
                kind=kind,
                role=role,
                x=(cast("float", rec.left) * 72.0) / width,
                y=(cast("float", rec.top) * 72.0) / height,
                width=(cast("float", rec.width) * 72.0) / width,
                height=(cast("float", rec.height) * 72.0) / height,
                slide_index=rec.slide_index,
                foreground=runs[0].color if runs else None,
                background=background,
                text=rec.text,
                font_size=font_size,
                font_size_unit="pt",
                font_size_norm=(font_size / height) if (font_size and height) else None,
                runs=runs,
                opacity=opacity,
                fill_kind=rec.fill_kind if rec.fill_kind != "unknown" else None,
                source="pptx",
            )
        )
    return out


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
        records=records,
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
