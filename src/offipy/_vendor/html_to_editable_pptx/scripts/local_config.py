"""local_config.py — 本机状态加载 + 首次 seed。

- `.config.local.toml`（gitignored）：用户偏好（fonts.auto_install / cleanup.default / audit.mode）
- `lessons-learned.md`（gitignored）：本地沉淀工作副本；首次缺失从
  `lessons-learned.md.example`（committed 模板）seed

可变数据一律落用户数据目录（OFFIPY_CONVERTER_DATA_DIR 优先），绝不写包内——
转换器 vendored 进 site-packages 后是只读的。offipy 拉起转换器时通过
OFFIPY_CONVERTER_DATA_DIR 注入；独立跑 convert.py 时回退平台用户目录。

文件缺失 / 解析失败 / key 缺 → 走 `_DEFAULTS`。永不抛。
"""
import os
import sys
from pathlib import Path

try:
    import tomllib  # py 3.11+
except ImportError:
    try:
        import tomli as tomllib  # py 3.10 的第三方等价实现
    except ImportError:
        tomllib = None

SKILL_ROOT = Path(__file__).resolve().parent.parent
LESSONS_TEMPLATE = SKILL_ROOT / "references" / "lessons-learned.md.example"


def _data_dir() -> Path:
    """可变数据目录：OFFIPY_CONVERTER_DATA_DIR 优先，否则平台用户目录。"""
    override = os.environ.get("OFFIPY_CONVERTER_DATA_DIR")
    if override:
        return Path(override)
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA")
        root = Path(base) if base else Path.home() / ".offipy"
    else:
        root = Path(os.environ.get("XDG_DATA_HOME", str(Path.home() / ".local" / "share")))
    return root / "offipy" / "converter"


CONFIG_PATH = _data_dir() / ".config.local.toml"
LESSONS_LOCAL = _data_dir() / "lessons-learned.md"

_DEFAULTS: dict = {
    "fonts": {"auto_install": "ask"},   # "yes" / "no" / "ask"
    "cleanup": {"default": "clean"},    # "clean" / "keep"
    "audit": {"mode": "ask"},           # "triage" / "page" / "manual" / "ask"
}


_warned_no_toml = False


def load() -> dict:
    global _warned_no_toml
    if not CONFIG_PATH.exists():
        return {k: {**v} for k, v in _DEFAULTS.items()}
    if tomllib is None:
        # 配置文件明明存在却静默走默认，用户会以为设置没生效且无从排查——必须提示
        if not _warned_no_toml:
            print(f"[config] 检测到 {CONFIG_PATH.name} 但当前 Python 缺 tomllib/tomli"
                  "（需 Python ≥ 3.11 或 pip install tomli）——本次配置被忽略，走默认值")
            _warned_no_toml = True
        return {k: {**v} for k, v in _DEFAULTS.items()}
    try:
        with CONFIG_PATH.open("rb") as f:
            user = tomllib.load(f) or {}
    except Exception as e:
        print(f"[config] 解析 {CONFIG_PATH.name} 失败（用默认）: {e}")
        return {k: {**v} for k, v in _DEFAULTS.items()}
    merged = {k: {**v} for k, v in _DEFAULTS.items()}
    for section, kvs in user.items():
        if section in merged and isinstance(kvs, dict):
            merged[section].update(kvs)
    return merged


def fonts_auto_install() -> str:
    return str(load().get("fonts", {}).get("auto_install", "ask")).lower()


def cleanup_default() -> str:
    return str(load().get("cleanup", {}).get("default", "clean")).lower()


def audit_mode() -> str:
    mode = str(load().get("audit", {}).get("mode", "ask")).lower().replace("-", "_")
    aliases = {
        "contact": "triage",
        "contact_sheet": "triage",
        "contact_first": "triage",
        "per_page": "page",
        "page_by_page": "page",
        "human": "manual",
    }
    mode = aliases.get(mode, mode)
    if mode not in {"triage", "page", "manual", "ask"}:
        print(f"[config] audit.mode={mode!r} 无效，回退 ask")
        return "ask"
    return mode


def seed_lessons_learned() -> None:
    """首次运行：本地 lessons-learned.md 缺失就从 .example 模板复制一份。

    模板上游会被覆盖，本地副本不会——agent 在本地副本上自由加 / 改 / 整理，
    打算上游的条目作者手动复制回 .example 模板再提交。
    """
    if LESSONS_LOCAL.exists() or not LESSONS_TEMPLATE.exists():
        return
    LESSONS_LOCAL.parent.mkdir(parents=True, exist_ok=True)
    LESSONS_LOCAL.write_text(LESSONS_TEMPLATE.read_text(encoding="utf-8"), encoding="utf-8")
    print(f"[lessons-learned] 首次运行：从模板 seed 本地工作副本 → {LESSONS_LOCAL.name}")
