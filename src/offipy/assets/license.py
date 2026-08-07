"""offipy.assets — license and provenance policy (rev1.2 §3.9).

V1 SPDX allowlist: MIT, ISC, CC0-1.0, CC-BY-4.0. `validate_provider_meta` is
the basic per-provider gate; `validate_manifest` is the strict vendored-source
gate that also requires source, source_commit and a matching count.
"""

from __future__ import annotations

from collections.abc import Mapping

from offipy.assets.model import AssetProviderMeta
from offipy.exceptions import InvalidArgumentError

ALLOWED_LICENSES = frozenset({"MIT", "ISC", "CC0-1.0", "CC-BY-4.0"})

_REQUIRED_MANIFEST_KEYS = ("license", "source", "source_commit", "count")


class LicensePolicy:
    """Validate license metadata against the allowlist."""

    def validate_provider_meta(self, meta: AssetProviderMeta) -> None:
        if meta.license not in ALLOWED_LICENSES:
            raise InvalidArgumentError(
                f"license {meta.license!r} not in allowlist {sorted(ALLOWED_LICENSES)}"
            )
        if meta.license == "CC-BY-4.0":
            if not meta.attribution or not meta.attribution.strip():
                raise InvalidArgumentError("CC-BY-4.0 requires non-empty attribution")
            if not meta.source_url or not meta.source_url.strip():
                raise InvalidArgumentError("CC-BY-4.0 requires a source_url")

    def validate_manifest(
        self, data: Mapping[str, object], *, actual_count: int | None = None
    ) -> None:
        missing = [k for k in _REQUIRED_MANIFEST_KEYS if k not in data]
        if missing:
            raise InvalidArgumentError(f"vendored manifest missing required keys {missing}")
        license_id = data["license"]
        if not isinstance(license_id, str) or license_id not in ALLOWED_LICENSES:
            raise InvalidArgumentError(f"vendored manifest license {license_id!r} not in allowlist")
        source = data["source"]
        if not isinstance(source, str) or not source.strip():
            raise InvalidArgumentError("vendored manifest source must be a non-empty URL")
        commit = data["source_commit"]
        if not isinstance(commit, str) or not commit.strip():
            raise InvalidArgumentError("vendored manifest source_commit must be non-empty")
        count = data["count"]
        if not isinstance(count, int) or isinstance(count, bool) or count < 0:
            raise InvalidArgumentError(
                f"vendored manifest count must be a non-negative int, got {count!r}"
            )
        if actual_count is not None and count != actual_count:
            raise InvalidArgumentError(f"vendored manifest count {count} != actual {actual_count}")
        if data.get("redistributable", True) is False:
            raise InvalidArgumentError("vendored manifest requires redistributable assets")
        if license_id == "CC-BY-4.0":
            attribution = data.get("attribution")
            if not isinstance(attribution, str) or not attribution.strip():
                raise InvalidArgumentError("vendored CC-BY-4.0 manifest requires attribution")
