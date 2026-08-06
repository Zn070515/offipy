"""艺术报告基线对比 v2：Finding 新增/解决/未变/变好/变坏/变化 + 维度 grade 变化。"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from typing import Literal

from .models import ArtFinding, ArtReport, ArtWarning, Grade

ChangeStatus = Literal["new", "resolved", "unchanged", "improved", "worsened", "changed"]


@dataclass
class FindingChange:
    rule_id: str
    dimension: str
    slide_index: int | None  # deck 级 finding（无页）为 None
    status: ChangeStatus
    before: ArtFinding | None = None
    after: ArtFinding | None = None


@dataclass
class GradeChange:
    dimension: str
    slide_index: int
    before: Grade | None
    after: Grade | None


@dataclass
class ArtReportDiff:
    before: ArtReport
    after: ArtReport
    changes: list[FindingChange] = field(default_factory=list)
    grade_changes: list[GradeChange] = field(default_factory=list)
    warnings: list[ArtWarning] = field(default_factory=list)

    @property
    def new_findings(self) -> list[FindingChange]:
        return [c for c in self.changes if c.status == "new"]

    @property
    def resolved_findings(self) -> list[FindingChange]:
        return [c for c in self.changes if c.status == "resolved"]


def _occurrence(f: ArtFinding) -> str:
    if f.primary:
        raw = "".join(sorted(r.element_id for r in f.related))
        raw += f.primary.kind + f.primary.role
    else:
        # 无 primary（deck 级 finding）：用 message+details 定身份，避免全 "none" 冲突
        raw = f.message + "".join(sorted(f"{k}:{v}" for k, v in f.details.items()))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]


def _stable_key(f: ArtFinding) -> tuple:
    return (f.rule_id, f.slide_index, f.primary.element_id if f.primary else None, _occurrence(f))


def _flatten(report: ArtReport) -> dict[tuple, ArtFinding]:
    flat: dict[tuple, ArtFinding] = {}
    for s in report.slides:
        for d in s.dimensions:
            for f in d.findings:
                flat[_stable_key(f)] = f
    for f in report.deck_findings:
        flat[_stable_key(f)] = f
    return flat


def _status(before: ArtFinding | None, after: ArtFinding | None) -> ChangeStatus:
    if before is None:
        return "new"
    if after is None:
        return "resolved"
    b_grade = _finding_grade(before)
    a_grade = _finding_grade(after)
    if a_grade is not None and b_grade is not None and a_grade != b_grade:
        return "improved" if a_grade < b_grade else "worsened"
    if (before.confidence, before.message, before.details) != (
        after.confidence,
        after.message,
        after.details,
    ):
        return "changed"
    if (before.evidence_sources, before.evidence_method) != (
        after.evidence_sources,
        after.evidence_method,
    ):
        return "changed"
    return "unchanged"


def _finding_grade(f: ArtFinding) -> int | None:
    return {"LOW": 1, "MID": 2, "HIGH": 3}.get(f.severity.name)


def _grade_changes(before: ArtReport, after: ArtReport) -> list[GradeChange]:
    b_grades = {(s.slide_index, d.dimension): d.grade for s in before.slides for d in s.dimensions}
    a_grades = {(s.slide_index, d.dimension): d.grade for s in after.slides for d in s.dimensions}
    out: list[GradeChange] = []
    for key in sorted(set(b_grades) | set(a_grades)):
        bg, ag = b_grades.get(key), a_grades.get(key)
        if bg != ag:
            out.append(GradeChange(dimension=key[1], slide_index=key[0], before=bg, after=ag))
    return out


def compare_reports(before: ArtReport, after: ArtReport) -> ArtReportDiff:
    warnings: list[ArtWarning] = []
    if before.schema_version != after.schema_version:
        warnings.append(
            ArtWarning(
                code="art.compare.schema_mismatch",
                message=f"schema {before.schema_version} → {after.schema_version}，对比仅建议性",
            )
        )
    if before.profile != after.profile:
        warnings.append(
            ArtWarning(
                code="art.compare.profile_mismatch",
                message=f"profile {before.profile} → {after.profile}，阈值不同",
            )
        )
    b_flat = _flatten(before)
    a_flat = _flatten(after)
    changes: list[FindingChange] = []
    for key, f in a_flat.items():
        b = b_flat.get(key)
        changes.append(
            FindingChange(
                f.rule_id,
                f.dimension,
                f.slide_index,
                _status(b, f),
                b,
                f,
            )
        )
    for key, f in b_flat.items():
        if key not in a_flat:
            changes.append(
                FindingChange(
                    f.rule_id,
                    f.dimension,
                    f.slide_index,
                    "resolved",
                    f,
                    None,
                )
            )
    # slide_index 可为 None（deck 级 finding）→ 用 0 占位避免混排 TypeError
    changes.sort(key=lambda c: (c.slide_index or 0, c.rule_id, c.status))
    return ArtReportDiff(
        before=before,
        after=after,
        changes=changes,
        grade_changes=_grade_changes(before, after),
        warnings=warnings,
    )
