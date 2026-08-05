> [中文](release.md)

# Release Guide

offipy is released via the **CI automated pipeline** (recommended) or **manual release** (fallback
for the first release or when CI is not configured). This document covers both, focusing on the
complete steps for pre-releasing 0.9.0a* to TestPyPI.

> The version number has a single source of truth: `__version__` in `src/offipy/__init__.py`. The
> tag `v<版本>` must match it (PEP 440, including `a`/`b`/`rc` pre-release suffixes). See the
> "Version Number Update Rules" section in the root `CLAUDE.md`.

---

## 0. Prerequisites (one-time setup)

| Prerequisite | Description |
|------|------|
| PyPI / TestPyPI account | Register accounts with the same name on [pypi.org](https://pypi.org) and [test.pypi.org](https://test.pypi.org) |
| Release method (choose one) | **① Trusted Publishing (OIDC, recommended)**: configure this GitHub repository `Zn070515/office-kit` as a trusted publisher on the PyPI/TestPyPI "Publishing" page (can be restricted to the workflow `release.yml` and the `testpypi`/`pypi` environments); **② API token**: for manual `twine upload`, via `~/.pypirc` or the environment variables `TWINE_USERNAME=__token__` + `TWINE_PASSWORD` |
| `twine` | Required for manual releases: `uv tool install twine` or `uvx twine` (CI uses `uvx twine` to avoid installing) |

The CI automated pipeline (`.github/workflows/release.yml`) runs when a `v*` tag is pushed; its
publish job relies on the OIDC trust configuration of TestPyPI/PyPI — **the publish job will fail
until configured**, in which case use the manual release below.

---

## 1. Local Quality Gates (run before release; only release when all pass)

```bash
export PYTHONIOENCODING=utf-8
uv run ruff check .
uv run ruff format --check .
uv run mypy src/offipy
uv run pytest tests -q          # COM/deck_render 无环境自动跳过
uv build
uvx twine check dist/*
```

Real-machine coverage (before merging to main) runs through `office-real` (self-hosted real
Office machine: Windows + real-machine Office COM + deck_render); locally you can run:
`uv run pytest tests -m "com or deck_render" -q`. After running, confirm
`tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` shows zero residue.

---

## 2. Pre-release to TestPyPI (0.9.0a*)

Pre-release versions go only to TestPyPI, not to the PyPI stable release. The `publish-testpypi`
job does this automatically in CI; the manual fallback is:

```bash
# 1) 构建 + 检查
uv build && uvx twine check dist/*

# 2) 上传 TestPyPI（token 方式；或先用 gh 配好 OIDC 后由 CI 发）
uvx twine upload --repository testpypi dist/*

# 3) 冒烟：从 TestPyPI 装到干净 venv，验证 import / 版本 / check 可跑
uv run python scripts/pypi_smoke.py --index https://test.pypi.org/simple --version 0.9.0a1
```

`pypi_smoke.py` exits non-zero if any verification point fails. Note that it does **not** assert
the exit code of `offipy check` — whether Chromium/Office are ready is a runtime environment
concern; the script only proves the package itself can be installed and run.

---

## 3. Stable Release (MAJOR.MINOR.PATCH)

Only stable releases go to the PyPI stable index (the CI `publish-pypi` job triggers automatically
for non-pre-release tags). **Hard gates:**

1. `office-real` real-machine tests pass (mandatory for release);
2. `pypi_smoke.py` verifies the 0.9.0a* smoke test from TestPyPI passes;
3. Manually confirm CHANGELOG / README / `__version__` / tag all agree, with no leftover old
   version numbers.

CI path: after `git push origin v<版本>`, `release.yml` completes everything automatically
(quality → office-real → gh-release → publish-testpypi → publish-pypi).

Manual fallback (when there is no OIDC): use the commands from Section 2 but replace
`--repository testpypi` with PyPI.

---

## 4. Post-Release Checklist

- [ ] `offipy check` all groups ✓ (`uv run offipy check`)
- [ ] `uv run python -c "import offipy; print(offipy.__version__)"` matches the tag
- [ ] The TestPyPI / PyPI page shows the corresponding version and files (sdist + wheel)
- [ ] The top-level CHANGELOG version matches the tag
