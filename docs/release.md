> [English](release.en.md)

# 发布手册

offipy 的发布走 **CI 自动管线**（推荐）或**手动发布**（首次/无 CI 配置时兜底）。
本文档覆盖两者，重点是 0.9.0a* 预发布到 TestPyPI 的完整步骤。

> 版本号单一来源：`src/offipy/__init__.py` 的 `__version__`。tag `v<版本>` 必须与它一致
> （PEP 440，含 `a`/`b`/`rc` 预发布后缀）。规则见根 `CLAUDE.md`「版本号更新规则」。

---

## 0. 发布前置（一次性的配置）

| 前置 | 说明 |
|------|------|
| PyPI / TestPyPI 账号 | 需在 [pypi.org](https://pypi.org) 和 [test.pypi.org](https://test.pypi.org) 注册同名账号 |
| 发布方式二选一 | **① Trusted Publishing（OIDC，推荐）**：在 PyPI/TestPyPI 的「Publishing」页把本 GitHub 仓库 `Zn070515/offipy` 配置为可信发布者（可限定 workflow `release.yml` 与环境 `testpypi`/`pypi`）；**② API token**：手动 `twine upload` 用，`~/.pypirc` 或环境变量 `TWINE_USERNAME=__token__` + `TWINE_PASSWORD` |
| `twine` | 手动发布需要：`uv tool install twine` 或 `uvx twine`（CI 里用 `uvx twine` 免安装） |

> **仓库改名提示**：若仓库更名，PyPI 上保存的旧仓库名**不会自动跟随**——Trusted Publisher
> 精确匹配仓库所有者 / 仓库名 / workflow 文件名 / environment，须用新名重配。

CI 自动管线（`.github/workflows/release.yml`）在推 `v*` tag 时运行，其发布 job 依赖
TestPyPI/PyPI 的 OIDC 信任配置——**未配置前发布 job 会失败**，此时走下面的手动发布。

---

## 1. 本地质量门禁（发布前必跑，全绿才发）

```bash
export PYTHONIOENCODING=utf-8
uv run ruff check .
uv run ruff format --check .
uv run mypy src/offipy
uv run pytest tests -q          # COM/deck_render 无环境自动跳过
uv build
uvx twine check dist/*
```

真机覆盖（合并到 main 前）走 `office-real`（自托管 Windows + Office 真机 COM + deck_render），
本地可跑：`uv run pytest tests -m "com or deck_render" -q`。跑完确认
`tasklist | grep -iE "POWERPNT|WINWORD|EXCEL"` 零残留。

---

## 2. 预发布到 TestPyPI（0.9.0a*）

预发布版本只上 TestPyPI，不上 PyPI 正式版。CI 里 `publish-testpypi` job 自动做；
手动兜底：

```bash
# 1) 构建 + 检查
uv build && uvx twine check dist/*

# 2) 上传 TestPyPI（token 方式；或先用 gh 配好 OIDC 后由 CI 发）
uvx twine upload --repository testpypi dist/*

# 3) 冒烟：从 TestPyPI 装到干净 venv，验证 import / 版本 / check 可跑
uv run python scripts/pypi_smoke.py --index https://test.pypi.org --version 0.9.0a1
```

`pypi_smoke.py` 任一验证点失败即非 0 退出。注意它**不断言 `offipy check` 的退出码**
——Chromium/Office 就绪是运行环境问题，只证明包本身可装可跑。

---

## 3. 正式版发布（MAJOR.MINOR.PATCH）

正式版才推 PyPI 正式版（CI `publish-pypi` 对非预发布 tag 自动触发）。**硬门禁：**

1. `office-real` 真机测试通过（发布必过）；
2. `pypi_smoke.py` 从 TestPyPI 验证 0.9.0a* 冒烟通过；
3. 手动确认 CHANGELOG / README / `__version__` / tag 四方一致，无残留旧版本号。

CI 路径：`git push origin v<版本>` 后由 `release.yml` 全自动完成
（quality → office-real → TestPyPI 发布 → TestPyPI 冒烟 → GitHub Release / PyPI 正式版）。
Release 不先于发布门禁——GitHub Release 与 PyPI 正式版都等 TestPyPI 精确安装冒烟通过后触发。

手动兜底（无 OIDC 时）：第 2 节的命令把 `--repository testpypi` 换成 PyPI 即可。

---

## 4. 发布后检查清单

- [ ] `offipy check` 全分组 ✓（`uv run offipy check`）
- [ ] `uv run python -c "import offipy; print(offipy.__version__)"` 与 tag 一致
- [ ] TestPyPI / PyPI 页面能看到对应版本与文件（sdist + wheel）
- [ ] CHANGELOG 顶层版本与 tag 一致
