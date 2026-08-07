"""A3 Task 9 — assets.json 事务性 provenance manifest。

Schema v1 冻结：{"schema": 1, "assets": [...]}；确定性 UTF-8、indent 2、EOF 换行。
provider 字段从 registry 的 AssetProviderMeta 序列化（ph→MIT、lu→ISC），
空报告 → 空 assets 列表。写入原子（tmp + os.replace）。
"""

import json

from offipy.assets.manifest import build_manifest_json, write_asset_manifest
from offipy.assets.model import AssetRef, AssetRequest
from offipy.assets.providers.icons import IconProvider
from offipy.assets.render import AssetUsageRecord, AssetUsageReport

SCHEMA = {"schema": 1, "assets": []}


def _record(provider_id, name, declaration_id, slide_index=1, placement="replace"):
    resolved = IconProvider(provider_id).resolve(AssetRequest(AssetRef(provider_id, "icon", name)))
    return AssetUsageRecord(
        declaration_id=declaration_id,
        slide_index=slide_index,
        request=f"asset://{provider_id}/icon/{name}",
        placement=placement,
        provider=resolved.provider_meta,
    )


class TestManifest:
    def test_empty_report_emits_empty_assets(self):
        text = build_manifest_json(AssetUsageReport(()))
        assert json.loads(text) == SCHEMA
        assert text.endswith("\n")

    def test_phosphor_record_matches_schema(self):
        record = _record("ph", "check", "asset-s01-001")
        text = build_manifest_json(AssetUsageReport((record,)))
        data = json.loads(text)
        assert data["schema"] == 1
        asset = data["assets"][0]
        assert asset["declaration_id"] == "asset-s01-001"
        assert asset["slide_index"] == 1
        assert asset["request"] == "asset://ph/icon/check"
        assert asset["placement"] == "replace"
        provider = asset["provider"]
        assert provider["id"] == "ph"
        assert provider["license"] == "MIT"
        assert provider["source_url"] and provider["source_commit"]
        assert provider["attribution"] is None
        assert provider["redistributable"] is True
        assert provider["first_party"] is False

    def test_lucide_record_isc_license(self):
        record = _record("lu", "settings", "asset-s02-001", slide_index=2, placement="background")
        data = json.loads(build_manifest_json(AssetUsageReport((record,))))
        assert data["assets"][0]["provider"]["license"] == "ISC"
        assert data["assets"][0]["placement"] == "background"

    def test_declaration_order_preserved(self):
        records = (
            _record("ph", "check", "asset-s01-001"),
            _record("lu", "settings", "asset-s02-001", slide_index=2),
            _record("ph", "gear", "asset-s02-002", slide_index=2),
        )
        data = json.loads(build_manifest_json(AssetUsageReport(records)))
        assert [a["declaration_id"] for a in data["assets"]] == [
            "asset-s01-001",
            "asset-s02-001",
            "asset-s02-002",
        ]

    def test_deterministic_output(self):
        records = (
            _record("ph", "check", "asset-s01-001"),
            _record("lu", "settings", "asset-s02-001", slide_index=2),
        )
        report = AssetUsageReport(records)
        assert build_manifest_json(report) == build_manifest_json(report)

    def test_write_asset_manifest_file_and_cleanup_tmp(self, tmp_path):
        p = tmp_path / "assets.json"
        write_asset_manifest(p, AssetUsageReport(()))
        assert (
            p.read_text(encoding="utf-8") == json.dumps(SCHEMA, ensure_ascii=False, indent=2) + "\n"
        )
        assert not list(tmp_path.glob(".*.tmp"))  # 原子写入后无残留 tmp
