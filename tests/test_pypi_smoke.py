"""pypi_smoke 下载 + 双重 sha256 比对 + 安装冒烟解析的纯单元测试（不触真实 TestPyPI / uv）。"""

import hashlib
import subprocess
import urllib.error

import pytest
from scripts.pypi_smoke import (
    _download_and_verify,
    _download_wheel,
    _fetch_json,
    _pick_wheel_url,
    _run,
    _verify_check_report,
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


# --- _run 的 check 参数透传 ---


def test_run_passes_check_param_through(monkeypatch):
    captured = {}

    def fake_run(cmd, **kwargs):
        captured.update(kwargs)
        return subprocess.CompletedProcess(cmd, 0)

    monkeypatch.setattr("scripts.pypi_smoke.subprocess.run", fake_run)
    _run(["echo", "hi"])
    assert captured["check"] is True
    _run(["echo", "hi"], check=False)
    assert captured["check"] is False


# --- _verify_check_report：offipy check 非零退出码仍解析合法 JSON ---


def _check_report(stdout: str, returncode: int = 1) -> subprocess.CompletedProcess:
    return subprocess.CompletedProcess(
        ["offipy", "check", "--profile", "all", "--json"],
        returncode,
        stdout=stdout,
    )


def test_verify_check_report_accepts_nonzero_returncode_with_valid_json():
    import json

    payload = {
        "version": _VERSION,
        "ok": False,
        "fails": 1,
        "warns": 0,
        "checks": [
            {
                "section": "浏览器",
                "name": "Chromium",
                "ok": False,
                "warn": False,
                "detail": "无法启动（headless）",
                "hint": "pip install playwright && playwright install chromium",
            }
        ],
    }
    # 真实 offipy check --json 输出是多行 pretty JSON（indent=2）——曾用
    # splitlines()[-1] 只取末行 '}' 解析失败，回归由 indent 版本防住。
    pretty = json.dumps(payload, indent=2, ensure_ascii=False)
    report = _verify_check_report(_check_report(pretty, returncode=1), _VERSION)
    assert report["version"] == _VERSION
    assert isinstance(report["checks"], list)
    assert report["ok"] is False


def test_verify_check_report_rejects_version_mismatch():
    import json

    r = _check_report(json.dumps({"version": "0.9.0", "checks": []}), returncode=0)
    with pytest.raises(SystemExit, match=r"!= '0\.9\.0a1'"):
        _verify_check_report(r, _VERSION)


def test_verify_check_report_rejects_missing_checks():
    import json

    r = _check_report(json.dumps({"version": _VERSION}))
    with pytest.raises(SystemExit, match="缺 checks"):
        _verify_check_report(r, _VERSION)


def test_verify_check_report_rejects_invalid_json():
    r = _check_report("usage: offipy [-h] ... not-json-at-all")
    with pytest.raises(SystemExit, match="不是合法 JSON"):
        _verify_check_report(r, _VERSION)
