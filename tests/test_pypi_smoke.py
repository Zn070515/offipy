"""pypi_smoke 下载 + 双重 sha256 比对的纯单元测试（不触真实 TestPyPI / uv）。"""

import hashlib
import urllib.error

import pytest
from scripts.pypi_smoke import (
    _download_and_verify,
    _download_wheel,
    _fetch_json,
    _pick_wheel_url,
)

_INDEX = "https://test.pypi.org"
_VERSION = "0.9.0a1"


def _wheel_meta(sha: str, upload: str, filename: str = "offipy-0.9.0a1-py3-none-any.whl") -> dict:
    return {
        "filename": filename,
        "packagetype": "bdist_wheel",
        "url": f"https://test.pypi.org/packages/fake/{filename}",
        "digests": {"sha256": sha},
        "upload_time_iso_8601": upload,
    }


# --- _pick_wheel_url ---


def test_pick_wheel_url_prefers_latest_wheel_and_skips_sdist():
    sha_old, sha_new = "a" * 64, "b" * 64
    data = {
        "urls": [
            {"filename": "offipy-0.9.0a1.tar.gz", "packagetype": "sdist"},
            _wheel_meta(sha_old, "2026-08-05T00:00:00Z"),
            _wheel_meta(sha_new, "2026-08-05T01:00:00Z"),
        ]
    }
    url, sha = _pick_wheel_url(data, _VERSION)
    assert url.endswith("offipy-0.9.0a1-py3-none-any.whl")
    assert sha == sha_new


def test_pick_wheel_url_raises_when_no_wheel():
    data = {"urls": [{"filename": "offipy-0.9.0a1.tar.gz", "packagetype": "sdist"}]}
    with pytest.raises(SystemExit, match="没有 wheel"):
        _pick_wheel_url(data, _VERSION)


def test_pick_wheel_url_raises_when_missing_sha256():
    meta = _wheel_meta("", "2026-08-05T00:00:00Z")
    meta["digests"] = {}
    with pytest.raises(SystemExit, match="缺 digests.sha256"):
        _pick_wheel_url({"urls": [meta]}, _VERSION)


# --- _fetch_json / _download_wheel 网络层 ---


class _FakeResp:
    def __init__(self, body: bytes):
        self._body = body

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def read(self):
        return self._body


def test_fetch_json_success(monkeypatch):
    monkeypatch.setattr(
        "urllib.request.urlopen",
        lambda url, timeout=60: _FakeResp(b'{"urls": []}'),
    )
    assert _fetch_json(f"{_INDEX}/pypi/offipy/{_VERSION}/json") == {"urls": []}


def test_fetch_json_http_error_raises_systemexit(monkeypatch):
    def boom(url, timeout=60):
        raise urllib.error.HTTPError(url, 404, "Not Found", None, None)

    monkeypatch.setattr("urllib.request.urlopen", boom)
    with pytest.raises(SystemExit, match="无法获取"):
        _fetch_json(f"{_INDEX}/pypi/offipy/{_VERSION}/json")


def test_download_wheel_http_error_raises_systemexit(monkeypatch, tmp_path):
    def boom(url, timeout=120):
        raise urllib.error.URLError("boom")

    monkeypatch.setattr("urllib.request.urlopen", boom)
    with pytest.raises(SystemExit, match="无法下载 wheel"):
        _download_wheel("https://test.pypi.org/packages/fake/offipy.whl", tmp_path / "offipy.whl")


# --- _download_and_verify：双重 sha256 比对 ---


def _stub_fetch_download(monkeypatch, payload: bytes, index_sha: str):
    meta = _wheel_meta(index_sha, "2026-08-05T00:00:00Z")
    monkeypatch.setattr(
        "scripts.pypi_smoke._fetch_json",
        lambda url: {"urls": [meta]},
    )
    monkeypatch.setattr(
        "scripts.pypi_smoke._download_wheel",
        lambda url, dest: dest.write_bytes(payload),
    )


def test_download_and_verify_success(tmp_path, monkeypatch):
    payload = b"offipy wheel payload"
    sha = hashlib.sha256(payload).hexdigest()
    _stub_fetch_download(monkeypatch, payload, sha)

    wheel = _download_and_verify(_INDEX, _VERSION, tmp_path, expected_sha256=sha)
    assert wheel.name == "offipy-0.9.0a1-py3-none-any.whl"
    assert wheel.read_bytes() == payload


def test_download_and_verify_fails_on_index_sha_mismatch(tmp_path, monkeypatch):
    payload = b"offipy wheel payload"
    _stub_fetch_download(monkeypatch, payload, index_sha="f" * 64)

    with pytest.raises(SystemExit, match="!= TestPyPI digests.sha256"):
        _download_and_verify(_INDEX, _VERSION, tmp_path, expected_sha256=None)


def test_download_and_verify_fails_on_expected_sha_mismatch(tmp_path, monkeypatch):
    payload = b"offipy wheel payload"
    sha = hashlib.sha256(payload).hexdigest()
    _stub_fetch_download(monkeypatch, payload, sha)

    with pytest.raises(SystemExit, match="!= 构建产物"):
        _download_and_verify(_INDEX, _VERSION, tmp_path, expected_sha256="e" * 64)
