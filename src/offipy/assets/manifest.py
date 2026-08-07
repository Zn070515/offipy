"""offipy.assets — transactional asset provenance manifest (assets.json).

Task 9: serialize a usage report into the audit dir's `assets.json`. Schema v1 is
frozen to `{"schema": 1, "assets": [...]}`; the writer is deterministic (UTF-8,
indent 2, trailing newline). Future readers may add fields; this writer emits
only the frozen schema. Writing is atomic (tmp + os.replace) so a failed render
never leaves a half-written manifest, and the deck's tmp-audit-dir → rename flow
means a failed render leaves the old final audit dir (and its assets.json) intact.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from offipy.assets.model import AssetProviderMeta
from offipy.assets.render import AssetUsageReport

_SCHEMA_VERSION = 1


def _provider_dict(provider: AssetProviderMeta) -> dict[str, object]:
    return {
        "id": provider.provider_id,
        "license": provider.license,
        "source_url": provider.source_url,
        "source_commit": provider.source_commit,
        "attribution": provider.attribution,
        "redistributable": provider.redistributable,
        "first_party": provider.first_party,
    }


def _record_dict(record) -> dict[str, object]:
    return {
        "declaration_id": record.declaration_id,
        "slide_index": record.slide_index,
        "request": record.request,
        "placement": record.placement,
        "provider": _provider_dict(record.provider),
    }


def build_manifest_json(report: AssetUsageReport) -> str:
    """Serialize a usage report to the frozen assets.json text (deterministic)."""
    payload: dict[str, object] = {
        "schema": _SCHEMA_VERSION,
        "assets": [_record_dict(r) for r in report.records],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2) + "\n"


def write_asset_manifest(path: str | Path, report: AssetUsageReport) -> None:
    """Atomically write assets.json into the audit dir.

    Writes to a sibling tmp then os.replace so a mid-write failure never leaves
    a truncated file; the deck renames the whole audit dir to its final name only
    after a successful render, so this only ever lands in tmp until commit.
    """
    p = Path(path)
    tmp = p.with_name(f".{p.name}.tmp")
    tmp.write_text(build_manifest_json(report), encoding="utf-8")
    os.replace(tmp, p)
