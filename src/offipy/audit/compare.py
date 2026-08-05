"""PPTX 基线回归对比：Shape 匹配链 + 聚合新增/已解决/变化的问题与形状变化。

惰性 import 硬约束：本模块模块级不触发 python-pptx（extract.py 内部才惰性
`from pptx import Presentation`）。

Shape 匹配链（每页按序，前一级命中即匹配）：
1. 同页相同 `shape_id`；
2. `name + shape_type`；
3. 归一化文本 hash + 中心距离（_MATCH_GEOM_TOL 内取最近）；
4. 图片内容 sha256（PICTURE 才可取得）。
仍未匹配 → 计入 unmatched（低置信，由调用方决定是否告警）。
"""

from __future__ import annotations

import hashlib
import math
from collections import defaultdict
from pathlib import Path
from typing import Any

from .extract import _PresentationExtract, _ShapeRecord
from .models import (
    AUDIT_SCHEMA_VERSION,
    AuditConfig,
    AuditFinding,
    AuditShapeRef,
    ChangedFinding,
    DiffShapeChange,
    PptxDiffReport,
)
from .rules import run_rules

_MOVE_TOL = 0.01  # 英寸，位置/尺寸变化判定容差
_MATCH_GEOM_TOL = 1.0  # 英寸，文本匹配时中心距离容差

# 类型别名：候选 (slide, shape_id) → 基线 (slide, shape_id)
_ShapeKey = tuple[int, int]
_KeyMap = dict[_ShapeKey, _ShapeKey]


def _sha256(path: str | Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _normalize_text(text: str) -> str:
    return " ".join(text.split())


def _shape_ref(rec: _ShapeRecord) -> AuditShapeRef:
    return AuditShapeRef(
        slide_index=rec.slide_index,
        shape_id=rec.shape_id,
        name=rec.name,
        shape_type=rec.shape_type,
        role=rec.role,
    )


def _center_dist(a: _ShapeRecord, b: _ShapeRecord) -> float:
    if (
        a.left is None
        or a.top is None
        or a.width is None
        or a.height is None
        or b.left is None
        or b.top is None
        or b.width is None
        or b.height is None
    ):
        return float("inf")
    ax = a.left + a.width / 2.0
    ay = a.top + a.height / 2.0
    bx = b.left + b.width / 2.0
    by = b.top + b.height / 2.0
    return math.hypot(ax - bx, ay - by)


def _match_slide(
    base: list[_ShapeRecord], cand: list[_ShapeRecord]
) -> tuple[dict[int, int], list[_ShapeRecord], list[_ShapeRecord]]:
    """单页 Shape 匹配 → (cand_id → base_id, 未匹配 base, 未匹配 cand)。"""
    matched_base: set[int] = set()
    matched_cand: set[int] = set()
    mapping: dict[int, int] = {}

    def claim(b_rec: _ShapeRecord, c_rec: _ShapeRecord) -> bool:
        if c_rec.shape_id in matched_cand:
            return False
        mapping[c_rec.shape_id] = b_rec.shape_id
        matched_base.add(b_rec.shape_id)
        matched_cand.add(c_rec.shape_id)
        return True

    # 1) 同 shape_id
    for b in base:
        if b.shape_id in matched_base:
            continue
        for c in cand:
            if c.shape_id not in matched_cand and c.shape_id == b.shape_id:
                claim(b, c)
                break

    # 2) name + shape_type
    for b in base:
        if b.shape_id in matched_base:
            continue
        for c in cand:
            if c.shape_id in matched_cand:
                continue
            if c.name == b.name and c.shape_type == b.shape_type:
                claim(b, c)
                break

    # 3) 归一化文本 hash + 几何邻近
    for b in base:
        if b.shape_id in matched_base:
            continue
        if not b.text.strip():
            continue
        best: _ShapeRecord | None = None
        best_d = float("inf")
        for c in cand:
            if c.shape_id in matched_cand:
                continue
            if _normalize_text(c.text) != _normalize_text(b.text):
                continue
            d = _center_dist(b, c)
            if d <= _MATCH_GEOM_TOL and d < best_d:
                best, best_d = c, d
        if best is not None:
            claim(b, best)

    # 4) 图片内容 hash
    for b in base:
        if b.shape_id in matched_base or not b.image_sha256:
            continue
        for c in cand:
            if c.shape_id in matched_cand:
                continue
            if c.image_sha256 and c.image_sha256 == b.image_sha256:
                claim(b, c)
                break

    return (
        mapping,
        [b for b in base if b.shape_id not in matched_base],
        [c for c in cand if c.shape_id not in matched_cand],
    )


def _change(rec: _ShapeRecord, kind: str, **extra: Any) -> DiffShapeChange:
    details: dict[str, Any] = {}
    if rec.left is not None:
        details["left_in"] = round(rec.left, 4)
    if rec.top is not None:
        details["top_in"] = round(rec.top, 4)
    if rec.width is not None:
        details["width_in"] = round(rec.width, 4)
    if rec.height is not None:
        details["height_in"] = round(rec.height, 4)
    if rec.text:
        details["text"] = rec.text
    details.update(extra)
    return DiffShapeChange(
        kind=kind,  # type: ignore[arg-type]
        slide_index=rec.slide_index,
        shape_id=rec.shape_id,
        name=rec.name,
        shape_type=rec.shape_type,
        details=details,
    )


def _record_geometry_changes(
    b: _ShapeRecord,
    c: _ShapeRecord,
    moved: list[DiffShapeChange],
    resized: list[DiffShapeChange],
    text_changes: list[DiffShapeChange],
) -> None:
    if (
        b.left is not None
        and b.top is not None
        and c.left is not None
        and c.top is not None
        and (abs(b.left - c.left) > _MOVE_TOL or abs(b.top - c.top) > _MOVE_TOL)
    ):
        moved.append(
            _change(
                c,
                "moved",
                old_left_in=round(b.left, 4),
                old_top_in=round(b.top, 4),
                new_left_in=round(c.left, 4),
                new_top_in=round(c.top, 4),
            )
        )
    if (
        b.width is not None
        and b.height is not None
        and c.width is not None
        and c.height is not None
        and (abs(b.width - c.width) > _MOVE_TOL or abs(b.height - c.height) > _MOVE_TOL)
    ):
        resized.append(
            _change(
                c,
                "resized",
                old_width_in=round(b.width, 4),
                old_height_in=round(b.height, 4),
                new_width_in=round(c.width, 4),
                new_height_in=round(c.height, 4),
            )
        )
    if b.text != c.text:
        text_changes.append(_change(c, "text", old_text=b.text, new_text=c.text))


def _base_finding_key(f: AuditFinding) -> tuple[Any, ...]:
    sec = (f.secondary.slide_index, f.secondary.shape_id) if f.secondary is not None else None
    return (f.rule_id, (f.primary.slide_index, f.primary.shape_id), sec)


def _cand_finding_key(f: AuditFinding, cand_to_base: _KeyMap) -> tuple[Any, ...] | None:
    """候选 Finding → 基线坐标空间的匹配键；主/次形状有任一未匹配则返回 None（新增）。"""
    pk = cand_to_base.get((f.primary.slide_index, f.primary.shape_id))
    if pk is None:
        return None
    if f.secondary is not None:
        sk = cand_to_base.get((f.secondary.slide_index, f.secondary.shape_id))
        if sk is None:
            return None
    else:
        sk = None
    return (f.rule_id, pk, sk)


def _diff_findings(
    base_findings: list[AuditFinding],
    cand_findings: list[AuditFinding],
    cand_to_base: _KeyMap,
) -> tuple[list[AuditFinding], list[AuditFinding], list[ChangedFinding]]:
    base_by_key: dict[tuple[Any, ...], list[AuditFinding]] = defaultdict(list)
    for f in base_findings:
        base_by_key[_base_finding_key(f)].append(f)

    added: list[AuditFinding] = []
    changed: list[ChangedFinding] = []
    for f in cand_findings:
        key = _cand_finding_key(f, cand_to_base)
        bucket = base_by_key.get(key) if key is not None else None
        if bucket is None or not bucket:
            added.append(f)
            continue
        bf = bucket.pop(0)
        if bf.severity != f.severity:
            changed.append(
                ChangedFinding(
                    rule_id=f.rule_id,
                    kind=f.kind,
                    old_severity=bf.severity,
                    new_severity=f.severity,
                    primary=f.primary,
                    secondary=f.secondary,
                    details=f.details,
                )
            )
    resolved = [f for fs in base_by_key.values() for f in fs]
    return added, resolved, changed


def build_diff(
    base_ext: _PresentationExtract,
    cand_ext: _PresentationExtract,
    *,
    baseline_path: str,
    candidate_path: str,
    offipy_version: str,
    audit_config: AuditConfig | None = None,
) -> PptxDiffReport:
    """对已提取的两个演示文稿做回归对比，产出 PptxDiffReport。"""
    cfg = audit_config or AuditConfig()
    base_recs = [r for slide in base_ext.slides for r in slide.shapes]
    cand_recs = [r for slide in cand_ext.slides for r in slide.shapes]
    base_findings, _ = run_rules(base_recs, base_ext.slide_size, cfg)
    cand_findings, _ = run_rules(cand_recs, cand_ext.slide_size, cfg)

    base_by_slide: dict[int, list[_ShapeRecord]] = defaultdict(list)
    cand_by_slide: dict[int, list[_ShapeRecord]] = defaultdict(list)
    for r in base_recs:
        base_by_slide[r.slide_index].append(r)
    for r in cand_recs:
        cand_by_slide[r.slide_index].append(r)

    cand_to_base: _KeyMap = {}
    added_shapes: list[DiffShapeChange] = []
    removed_shapes: list[DiffShapeChange] = []
    moved_shapes: list[DiffShapeChange] = []
    resized_shapes: list[DiffShapeChange] = []
    text_changes: list[DiffShapeChange] = []
    unmatched_base: list[_ShapeRecord] = []
    unmatched_cand: list[_ShapeRecord] = []

    shared = min(len(base_ext.slides), len(cand_ext.slides))
    for slide_idx in range(1, shared + 1):
        b_list = base_by_slide.get(slide_idx, [])
        c_list = cand_by_slide.get(slide_idx, [])
        mapping, ub, uc = _match_slide(b_list, c_list)
        for c_id, b_id in mapping.items():
            cand_to_base[(slide_idx, c_id)] = (slide_idx, b_id)
        b_by_id = {r.shape_id: r for r in b_list}
        c_by_id = {r.shape_id: r for r in c_list}
        for r in ub:
            removed_shapes.append(_change(r, "removed"))
        for r in uc:
            added_shapes.append(_change(r, "added"))
        for c_id, b_id in mapping.items():
            _record_geometry_changes(
                b_by_id[b_id], c_by_id[c_id], moved_shapes, resized_shapes, text_changes
            )
        unmatched_base.extend(ub)
        unmatched_cand.extend(uc)

    for slide_idx in range(shared + 1, len(cand_ext.slides) + 1):
        for r in cand_by_slide.get(slide_idx, []):
            added_shapes.append(_change(r, "added"))
    for slide_idx in range(shared + 1, len(base_ext.slides) + 1):
        for r in base_by_slide.get(slide_idx, []):
            removed_shapes.append(_change(r, "removed"))

    added_findings, resolved_findings, changed_findings = _diff_findings(
        base_findings, cand_findings, cand_to_base
    )

    return PptxDiffReport(
        schema_version=AUDIT_SCHEMA_VERSION,
        offipy_version=offipy_version,
        baseline_path=baseline_path,
        candidate_path=candidate_path,
        baseline_sha256=_sha256(baseline_path),
        candidate_sha256=_sha256(candidate_path),
        baseline_slide_count=len(base_ext.slides),
        candidate_slide_count=len(cand_ext.slides),
        baseline_findings=base_findings,
        candidate_findings=cand_findings,
        added_findings=added_findings,
        resolved_findings=resolved_findings,
        changed_findings=changed_findings,
        added_shapes=added_shapes,
        removed_shapes=removed_shapes,
        moved_shapes=moved_shapes,
        resized_shapes=resized_shapes,
        text_changes=text_changes,
        unmatched_baseline=[_shape_ref(r) for r in unmatched_base],
        unmatched_candidate=[_shape_ref(r) for r in unmatched_cand],
        warnings=[*base_ext.warnings, *cand_ext.warnings],
    )
