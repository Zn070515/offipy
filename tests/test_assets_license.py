"""A2 Task 4 — license and provenance policy."""

import json
from pathlib import Path

import pytest

from offipy.assets import AssetProviderMeta
from offipy.assets.license import ALLOWED_LICENSES, LicensePolicy
from offipy.assets.registry import AssetRegistry
from offipy.exceptions import InvalidArgumentError

_POLICY = LicensePolicy()

_REPO_ROOT = Path(__file__).resolve().parents[1]
_ICONS_DIR = _REPO_ROOT / "src" / "offipy" / "assets" / "icons"
_MANIFEST = _ICONS_DIR / "manifest.json"
_NOTICES = _REPO_ROOT / "THIRD_PARTY_NOTICES.md"


def _meta(**kw) -> AssetProviderMeta:
    defaults = {
        "provider_id": "ph",
        "license": "ISC",
        "source_url": None,
        "source_commit": None,
        "attribution": None,
        "redistributable": True,
        "first_party": False,
    }
    defaults.update(kw)
    return AssetProviderMeta(**defaults)


# ---------------------------------------------------------------------------
# validate_provider_meta
# ---------------------------------------------------------------------------


class TestValidateProviderMeta:
    def test_allowlist_accepted(self):
        assert frozenset({"MIT", "ISC", "CC0-1.0", "CC-BY-4.0"}) == ALLOWED_LICENSES
        for lic in ("MIT", "ISC", "CC0-1.0"):
            _POLICY.validate_provider_meta(_meta(license=lic))
        _POLICY.validate_provider_meta(
            _meta(license="CC-BY-4.0", source_url="https://x", attribution="A")
        )

    def test_unknown_license_fails(self):
        for lic in ("GPL-3.0", "Apache-2.0", ""):
            with pytest.raises(InvalidArgumentError):
                _POLICY.validate_provider_meta(_meta(license=lic))

    def test_cc_by_requires_attribution_and_source(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_provider_meta(_meta(license="CC-BY-4.0"))
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_provider_meta(_meta(license="CC-BY-4.0", source_url="https://x"))
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_provider_meta(_meta(license="CC-BY-4.0", attribution="A"))

    def test_first_party_may_lack_commit_but_identifies_license(self):
        _POLICY.validate_provider_meta(_meta(first_party=True, license="MIT", source_commit=None))


class _FakeProvider:
    """Minimal AssetProvider whose provider_meta can carry a chosen license."""

    def __init__(self, provider_id: str, license_: str) -> None:
        self.provider_id = provider_id
        self.kinds = frozenset({"icon"})
        self.provider_meta = _meta(provider_id=provider_id, license=license_)


class TestRegistryRuntimeEnforcement:
    def test_disallowed_license_blocked_at_register(self):
        reg = AssetRegistry()
        with pytest.raises(InvalidArgumentError, match="license"):
            reg.register(_FakeProvider("badlic", "GPL-3.0"))

    def test_allowlisted_license_registers(self):
        reg = AssetRegistry()
        reg.register(_FakeProvider("goodlic", "MIT"))
        assert reg.provider("goodlic").provider_id == "goodlic"

    def test_remote_third_party_may_lack_commit(self):
        # future/remote provider metadata may be redistributable=False without
        # a commit; the basic gate still accepts, vendored gate rejects.
        _POLICY.validate_provider_meta(
            _meta(redistributable=False, license="MIT", source_url="https://x", source_commit=None)
        )


# ---------------------------------------------------------------------------
# validate_manifest
# ---------------------------------------------------------------------------


def _manifest(**kw) -> dict[str, object]:
    base: dict[str, object] = {
        "license": "MIT",
        "source": "https://codeload.github.com/x/y/tar.gz/abc",
        "source_commit": "abc",
        "count": 10,
    }
    base.update(kw)
    return base


class TestValidateManifest:
    def test_valid_manifest(self):
        _POLICY.validate_manifest(_manifest(), actual_count=10)

    def test_count_mismatch_fails(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(), actual_count=11)

    def test_count_negative_fails(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(count=-1))

    def test_count_non_int_fails(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(count="10"))

    def test_missing_required_keys(self):
        for missing in ("license", "source", "source_commit", "count"):
            data = _manifest()
            del data[missing]
            with pytest.raises(InvalidArgumentError):
                _POLICY.validate_manifest(data)

    def test_empty_source_fails(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(source=""))

    def test_empty_commit_fails(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(source_commit=""))

    def test_unknown_license_fails(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(license="Apache-2.0"))

    def test_non_redistributable_fails_vendored(self):
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(redistributable=False))

    def test_cc_by_manifest_requires_attribution(self):
        _POLICY.validate_manifest(_manifest(license="CC-BY-4.0", attribution="Author"))
        with pytest.raises(InvalidArgumentError):
            _POLICY.validate_manifest(_manifest(license="CC-BY-4.0"))


# ---------------------------------------------------------------------------
# existing icons manifest (§5.3)
# ---------------------------------------------------------------------------


class TestExistingIconsManifest:
    def _manifest_data(self):
        return json.loads(_MANIFEST.read_text(encoding="utf-8"))

    def test_phosphor_count_matches_svg(self):
        data = self._manifest_data()
        svg_count = len(list((_ICONS_DIR / "phosphor").glob("*.svg")))
        assert data["phosphor"]["count"] == svg_count

    def test_lucide_count_matches_svg(self):
        data = self._manifest_data()
        svg_count = len(list((_ICONS_DIR / "lucide").glob("*.svg")))
        assert data["lucide"]["count"] == svg_count

    def test_licenses_accepted(self):
        data = self._manifest_data()
        for provider_id, entry in data.items():
            _POLICY.validate_provider_meta(
                AssetProviderMeta(
                    provider_id=provider_id,
                    license=entry["license"],
                    source_url=entry["source"],
                    source_commit=entry["commit"],
                    attribution=None,
                    redistributable=True,
                )
            )

    def test_source_commit_map_without_loss(self):
        data = self._manifest_data()
        for provider_id, entry in data.items():
            meta = AssetProviderMeta(
                provider_id=provider_id,
                license=entry["license"],
                source_url=entry["source"],
                source_commit=entry["commit"],
                attribution=None,
                redistributable=True,
            )
            assert meta.source_url == entry["source"]
            assert meta.source_commit == entry["commit"]


# ---------------------------------------------------------------------------
# THIRD_PARTY_NOTICES gate (§5.4)
# ---------------------------------------------------------------------------


class TestThirdPartyNoticesGate:
    def test_vendored_providers_mentioned(self):
        text = _NOTICES.read_text(encoding="utf-8")
        data = json.loads(_MANIFEST.read_text(encoding="utf-8"))
        for provider_id, entry in data.items():
            assert entry["license"] in text, f"license for {provider_id} missing"
            assert provider_id.capitalize() in text, f"{provider_id} name missing"
