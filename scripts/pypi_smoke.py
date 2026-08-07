"""TestPyPI 精确安装门禁（P0-5）：JSON API 精确下载 wheel → 双重 sha256 比对 → 干净 venv 安装冒烟。

任一断言失败 → 非 0 退出。验证点：
- GET {index}/pypi/offipy/<version>/json（TestPyPI JSON API，精确到版本）
- 从 urls 精确选择 .whl，下载后双重 sha256 比对：
    1) TestPyPI 索引声明的 digests.sha256
    2) --expected-sha256（构建 artifact 的 sha，门禁用）
  两者都必须等于本地下载 sha——证明「TestPyPI 上的就是本轮构建产物」
  （防「旧同版本被 skip」后拿旧文件蒙混）
- 新建干净 venv：装下载的 wheel，[deck,mcp] extras 运行时依赖从正式 PyPI 解析
  （Linux 上 pywin32 因 platform_system marker 自动跳过）
- `import offipy` / __version__ == --version / `offipy --help` /
  `offipy check --profile all --json`（合法 JSON 且 version 匹配）/
  `offipy mcp --help`

CLI:
    --version 必填；--index 为 JSON API base，默认 https://test.pypi.org，
    实际请求 {index}/pypi/offipy/<version>/json；
    --expected-sha256 给了才比对（构建 artifact 的 sha，门禁用）。

注：不断言 check 的退出码——Chromium/Office 是否就绪是运行环境问题，
不是打包问题；这里只证明「TestPyPI 上的 wheel == 构建产物，且可装可跑」。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

_DEFAULT_INDEX = "https://test.pypi.org"


def _venv_python(venv: Path) -> str:
    return str(venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python"))


def _venv_script(venv: Path, name: str) -> str:
    return str(venv / ("Scripts" if os.name == "nt" else "bin") / name)


def _run(
    cmd: list[str],
    *,
    check: bool = True,
) -> subprocess.CompletedProcess:
    # 子进程统一 UTF-8：Windows 默认 GBK 会把中文输出打乱/解码失败。
    # check=False 用于「退出码不代表门禁结果」的命令（如 offipy check——
    # Chromium/Office 是否就绪是运行环境问题，非打包问题）。
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    return subprocess.run(
        cmd,
        check=check,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _fetch_json(url: str) -> dict:
    """GET JSON；HTTP/URL/网络错误一律 SystemExit（门禁脚本不做半开）。"""
    try:
        with urllib.request.urlopen(url, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.HTTPError, urllib.error.URLError, OSError) as exc:
        raise SystemExit(f"[pypi-smoke] FAIL: 无法获取 {url}: {exc}") from exc


def _download_wheel(url: str, dest: Path) -> None:
    try:
        with urllib.request.urlopen(url, timeout=120) as resp, open(dest, "wb") as fh:
            shutil.copyfileobj(resp, fh)
    except (urllib.error.HTTPError, urllib.error.URLError, OSError) as exc:
        raise SystemExit(f"[pypi-smoke] FAIL: 无法下载 wheel {url}: {exc}") from exc


def _pick_wheel_url(data: dict, version: str) -> tuple[str, str]:
    """从 TestPyPI JSON 的 urls 里选 .whl，多个按上传时间取最新；返回 (url, digests.sha256)。"""
    wheels = [u for u in data.get("urls", []) if u.get("filename", "").endswith(".whl")]
    if not wheels:
        raise SystemExit(
            f"[pypi-smoke] FAIL: offipy=={version} JSON 里没有 wheel"
            f"（urls 共 {len(data.get('urls', []))} 个文件）"
        )
    wheels.sort(key=lambda u: u.get("upload_time_iso_8601", ""), reverse=True)
    chosen = wheels[0]
    sha = chosen.get("digests", {}).get("sha256")
    if not sha:
        raise SystemExit(f"[pypi-smoke] FAIL: {chosen['filename']} 缺 digests.sha256")
    return chosen["url"], sha


def _download_and_verify(
    index: str,
    version: str,
    download_dir: Path,
    expected_sha256: str | None,
) -> Path:
    """从 index JSON API 精确下载 wheel 并做双重 sha256 比对，返回 wheel 路径。

    比对 1：本地下载 sha == TestPyPI 索引声明的 digests.sha256；
    比对 2：若给了 expected_sha256（构建 artifact），本地 sha == expected。
    """
    json_url = f"{index}/pypi/offipy/{version}/json"
    print(f"[pypi-smoke] GET {json_url}（TestPyPI JSON API 精确查询）...")
    data = _fetch_json(json_url)
    wheel_url, index_sha = _pick_wheel_url(data, version)
    wheel = download_dir / Path(wheel_url.split("/")[-1])
    print(f"[pypi-smoke] 下载 {wheel.name} ...")
    _download_wheel(wheel_url, wheel)

    sha = _sha256(wheel)
    if sha != index_sha:
        raise SystemExit(
            f"[pypi-smoke] FAIL: 本地下载 sha256 {sha} != TestPyPI digests.sha256 "
            f"{index_sha}（索引声明与下载内容不符）"
        )
    print("[pypi-smoke] sha256 == TestPyPI digests.sha256 ✓")

    if expected_sha256:
        if sha != expected_sha256.lower():
            raise SystemExit(
                f"[pypi-smoke] FAIL: wheel sha256 {sha} != 构建产物 "
                f"{expected_sha256}（TestPyPI 上不是本轮构建内容）"
            )
        print("[pypi-smoke] sha256 == 构建产物 ✓")
    else:
        print(f"[pypi-smoke] sha256 = {sha}（未给 --expected-sha256，跳过构建产物比对）")
    return wheel


def _verify_check_report(r: subprocess.CompletedProcess, version: str) -> dict:
    """解析 `offipy check --json` 的 stdout 并核对 version/checks。

    不断言退出码：`offipy check` 返回 1 只代表运行环境某项未就绪
    （如缺 Chromium），仍是合法 JSON——smoke 的目标是证明包可装可跑。
    """
    raw = r.stdout.strip()
    try:
        report = json.loads(raw)  # offipy check --json 输出整段多行 pretty JSON
    except json.JSONDecodeError:
        # 兜底：若 stdout 前面混入非 JSON 行，JSON 在末行（紧凑单行场景）
        try:
            report = json.loads(raw.splitlines()[-1])
        except (json.JSONDecodeError, IndexError) as exc:
            raise SystemExit(
                f"[pypi-smoke] FAIL: offipy check --json 输出不是合法 JSON: {r.stdout!r}"
            ) from exc
    if report.get("version") != version:
        raise SystemExit(
            f"[pypi-smoke] FAIL: check JSON version = {report.get('version')!r} != {version!r}"
        )
    if not isinstance(report.get("checks"), list):
        raise SystemExit("[pypi-smoke] FAIL: check JSON 缺 checks 数组")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="TestPyPI 精确安装门禁。")
    parser.add_argument("--version", required=True, help="要下载安装验证的版本，如 0.9.0a1")
    parser.add_argument(
        "--index",
        default=_DEFAULT_INDEX,
        help=f"JSON API base（默认 {_DEFAULT_INDEX}），请求 {{index}}/pypi/offipy/<version>/json",
    )
    parser.add_argument(
        "--expected-sha256",
        help="构建 artifact 的 wheel sha256；给了则断言下载产物同一内容",
    )
    args = parser.parse_args()
    print(f"[pypi-smoke] index = {args.index}, version = {args.version}")

    tmp = Path(tempfile.mkdtemp(prefix="offipy-pypi-smoke-"))
    try:
        venv = tmp / "venv"
        download = tmp / "download"
        download.mkdir()
        print("[pypi-smoke] 创建干净 venv（用于解析 wheel 平台 + 安装）...")
        _run(["uv", "venv", str(venv)])
        py = _venv_python(venv)

        wheel = _download_and_verify(args.index, args.version, download, args.expected_sha256)

        print("[pypi-smoke] 干净 venv：装 wheel + PyPI 解析 [deck,mcp] extras ...")
        _run(
            [
                "uv",
                "pip",
                "install",
                "--python",
                py,
                "--index-url",
                "https://pypi.org/simple",
                f"{wheel}[deck,mcp]",
            ]
        )
        script = _venv_script(venv, "offipy")

        print("[pypi-smoke] import offipy + __version__ 匹配 ...")
        r = _run([py, "-c", "import offipy; print(offipy.__version__)"])
        if r.stdout.strip() != args.version:
            raise SystemExit(
                f"[pypi-smoke] FAIL: import offipy.__version__ = {r.stdout.strip()!r} "
                f"!= --version {args.version!r}"
            )
        print(f"[pypi-smoke] __version__ == {args.version} ✓")

        print("[pypi-smoke] offipy --help ...")
        r = _run([script, "--help"])
        if "usage" not in r.stdout.lower():
            raise SystemExit("[pypi-smoke] FAIL: offipy --help 无 usage 输出")

        print("[pypi-smoke] offipy check --profile all --json ...")
        # 不断言退出码：Chromium/Office 是否就绪是运行环境问题，非打包问题；
        # 只要 stdout 是合法 JSON 且 version 匹配即可（环境就绪由 office-real 承担）。
        r = _run([script, "check", "--profile", "all", "--json"], check=False)
        report = _verify_check_report(r, args.version)
        print(
            f"[pypi-smoke] check JSON version 匹配，checks = {len(report['checks'])} 项 "
            f"（rc={r.returncode} 不断言，环境就绪由 office-real 承担）"
        )

        print("[pypi-smoke] offipy mcp --help ...")
        r = _run([script, "mcp", "--help"])
        if "usage" not in r.stdout.lower():
            raise SystemExit("[pypi-smoke] FAIL: offipy mcp --help 无 usage 输出")

        print("[pypi-smoke] 资产 API 解析（ph/lu/procedural/primitives）...")
        # 证明 wheel 内含 vendored 图标/纹理/图元模块与 SVG 文件，且 import offipy.assets
        # 保持纯标准库（解析不依赖 python-pptx/PIL/playwright）
        asset_code = (
            "import json\n"
            "from offipy.assets import get_default_registry\n"
            "r = get_default_registry()\n"
            "names = [\n"
            "    r.resolve('asset://ph/icon/check').meta.ref.name,\n"
            "    r.resolve('asset://lu/icon/settings').meta.ref.name,\n"
            "    r.resolve('asset://procedural/pattern/wave').meta.ref.name,\n"
            "    r.resolve('asset://primitives/primitive/browser-mockup').meta.ref.name,\n"
            "]\n"
            "print(json.dumps(names))\n"
        )
        r = _run([py, "-c", asset_code])
        try:
            parsed = json.loads(r.stdout)
        except json.JSONDecodeError as exc:
            raise SystemExit(f"[pypi-smoke] FAIL: 资产 API 输出非 JSON: {r.stdout!r}") from exc
        if parsed != ["check", "settings", "wave", "browser-mockup"]:
            raise SystemExit(f"[pypi-smoke] FAIL: 资产解析结果 {parsed!r} 不符预期")
        print("[pypi-smoke] 资产 API 解析 4/4 ✓")

        print(f"[pypi-smoke] OK — {args.version} 从 TestPyPI 精确下载 + 干净安装冒烟通过")
        return 0
    finally:
        print(f"[pypi-smoke] 清理 {tmp}")
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
