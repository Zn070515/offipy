"""audit 公共模型：Severity 序/序列化、AuditConfig 默认值、JSON 安全、空报告。"""

import json

import pytest

from offipy.audit import (
    ALL_RULE_IDS,
    AUDIT_SCHEMA_VERSION,
    RULE_BOUNDS_PARTIAL,
    RULE_TEXT_FIT_HORIZONTAL,
    AuditConfig,
    AuditFinding,
    AuditShapeRef,
    AuditWarning,
    PptxAuditReport,
    Severity,
    SuppressedFinding,
)
from offipy.audit.models import JsonValue

# ---------------------------------------------------------------- Severity


@pytest.mark.parametrize(
    "sev,name",
    [(Severity.LOW, "LOW"), (Severity.MID, "MID"), (Severity.HIGH, "HIGH")],
)
def test_severity_order_and_name(sev, name):
    assert sev.value in (1, 2, 3)
    assert sev.name == name


def test_severity_ordered_by_integer_value():
    assert Severity.LOW < Severity.MID < Severity.HIGH
    assert Severity.LOW.value == 1
    assert Severity.HIGH.value == 3


def test_all_rule_ids_unique():
    assert len(set(ALL_RULE_IDS)) == len(ALL_RULE_IDS)
    assert RULE_BOUNDS_PARTIAL == "geometry.bounds.partial"
    assert RULE_TEXT_FIT_HORIZONTAL == "text.fit.horizontal"


# ---------------------------------------------------------------- AuditConfig


def test_audit_config_defaults():
    c = AuditConfig()
    assert c.safe_margin_in == 0.2
    assert c.bounds_tolerance_in == 0.01
    assert c.ignore_full_bleed_shapes is True
    assert c.ignore_repeated_decorations is True
    assert c.ignore_page_numbers is True
    assert c.ignore_headers_footers is True
    assert c.ignored_shapes is None
    assert c.ignored_regions is None


def test_audit_config_to_dict_json_safe():
    c = AuditConfig(
        ignored_shapes={(2, 15), (1, 7)},
        ignored_regions=[(0.5, 0.5, 2.0, 1.0)],
    )
    d = c.to_dict()
    json.dumps(d)  # 可序列化即 JSON 安全
    assert d["ignored_shapes"] == [(1, 7), (2, 15)]  # set 转排序列表
    assert d["ignored_regions"] == [(0.5, 0.5, 2.0, 1.0)]


# ---------------------------------------------------------------- Finding JSON


def _ref() -> AuditShapeRef:
    return AuditShapeRef(
        slide_index=2,
        shape_id=11,
        name="t1",
        shape_type="TEXT_BOX",
        role="content",
    )


def test_finding_to_dict_json_safe():
    f = AuditFinding(
        rule_id=RULE_BOUNDS_PARTIAL,
        kind="bounds",
        severity=Severity.MID,
        message="部分越界",
        primary=_ref(),
        details={"overflow_in": 0.3, "side": "right"},
    )
    d = f.to_dict()
    assert d["severity"] == "MID"
    assert d["rule_id"] == RULE_BOUNDS_PARTIAL
    assert d["primary"]["slide_index"] == 2
    assert d["details"]["overflow_in"] == 0.3
    assert "secondary" not in d
    json.dumps(d, ensure_ascii=False)


def test_finding_with_secondary():
    f = AuditFinding(
        rule_id=RULE_BOUNDS_PARTIAL,
        kind="overlap",
        severity=Severity.HIGH,
        message="完全覆盖",
        primary=_ref(),
        secondary=_ref(),
    )
    assert f.to_dict()["secondary"]["shape_id"] == 11
    assert f.confidence == 1.0


def test_suppressed_finding_to_dict():
    f = AuditFinding(
        rule_id=RULE_BOUNDS_PARTIAL,
        kind="margin",
        severity=Severity.LOW,
        message="贴边",
        primary=_ref(),
    )
    s = SuppressedFinding(finding=f, reason="page_number")
    d = s.to_dict()
    assert d["reason"] == "page_number"
    assert d["finding"]["severity"] == "LOW"


def test_warning_to_dict_json_safe():
    w = AuditWarning(
        slide_index=3,
        shape_id=9,
        code="rot.non_axis",
        message="旋转无法精确转换",
    )
    d = w.to_dict()
    assert d == {
        "slide_index": 3,
        "shape_id": 9,
        "code": "rot.non_axis",
        "message": "旋转无法精确转换",
    }
    json.dumps(d, ensure_ascii=False)


def test_warning_none_fields():
    w = AuditWarning(slide_index=None, shape_id=None, code="parse", message="x")
    assert w.to_dict()["slide_index"] is None


# ---------------------------------------------------------------- 空报告


def test_empty_report_max_severity_none():
    r = PptxAuditReport(
        schema_version=AUDIT_SCHEMA_VERSION,
        offipy_version="0.0.0",
        path="x.pptx",
        source_sha256="ab" * 32,
        slide_size=(13.333, 7.5),
        slide_count=1,
        config=AuditConfig(),
    )
    assert r.max_severity is None
    d = r.to_dict()
    assert d["max_severity"] is None
    assert d["findings"] == []
    assert d["suppressed"] == []
    assert d["warnings"] == []
    json.dumps(d, ensure_ascii=False)


def test_report_max_severity_takes_max():
    lo = AuditFinding(
        rule_id=RULE_BOUNDS_PARTIAL,
        kind="margin",
        severity=Severity.LOW,
        message="低",
        primary=_ref(),
    )
    hi = AuditFinding(
        rule_id=RULE_BOUNDS_PARTIAL,
        kind="bounds",
        severity=Severity.HIGH,
        message="高",
        primary=_ref(),
    )
    r = PptxAuditReport(
        schema_version=AUDIT_SCHEMA_VERSION,
        offipy_version="0.0.0",
        path="x.pptx",
        source_sha256="ab" * 32,
        slide_size=(13.333, 7.5),
        slide_count=1,
        config=AuditConfig(),
        findings=[lo, hi],
    )
    assert r.max_severity is Severity.HIGH
    assert r.to_dict()["max_severity"] == "HIGH"


def test_report_to_json_roundtrip():
    r = PptxAuditReport(
        schema_version=AUDIT_SCHEMA_VERSION,
        offipy_version="0.0.0",
        path="x.pptx",
        source_sha256="ab" * 32,
        slide_size=(13.333, 7.5),
        slide_count=1,
        config=AuditConfig(ignored_shapes={(1, 3)}),
    )
    data = json.loads(r.to_json())
    assert data["config"]["ignored_shapes"] == [[1, 3]]
    assert data["slide_size"] == [13.333, 7.5]


# ---------------------------------------------------------------- 门禁：不触发 python-pptx


def test_import_audit_models_does_not_load_pptx(monkeypatch):
    import builtins

    orig_import = builtins.__import__

    def guarded(name, *args, **kwargs):
        if name == "pptx" or name.startswith("pptx."):
            raise AssertionError("pptx 不应被 audit.models 导入")
        return orig_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", guarded)
    from offipy.audit.models import Severity as S  # noqa: F401

    assert S.HIGH == 3


def test_json_value_alias_accepts_common_types():
    v: JsonValue = {"a": [1, "x", None, True]}
    assert isinstance(v, dict)
