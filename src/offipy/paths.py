"""用户数据目录：可变数据一律写用户目录，绝不写进 site-packages / 包目录。

- server token（批次 3 用）、.offipy.log、转换器的 lessons-learned 落盘
  全部经这里定位，保证安装在只读 site-packages 后也能正常工作。
- converter_data_dir() 优先读 OFFIPY_CONVERTER_DATA_DIR 环境变量，
  供测试/CI 注入临时目录，也便于用户显式迁移数据。
"""

import os
import sys
from pathlib import Path

_APP_DIRNAME = "offipy"
_CONVERTER_DATA_ENV = "OFFIPY_CONVERTER_DATA_DIR"


def user_data_dir() -> Path:
    """跨平台用户数据目录。

    - Windows: %LOCALAPPDATA%\\offipy（缺 LOCALAPPDATA 时退回 ~/.offipy）
    - macOS:   ~/Library/Application Support/offipy
    - Linux:   $XDG_DATA_HOME/offipy（缺省 ~/.local/share/offipy）
    """
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA")
        if base:
            return Path(base) / _APP_DIRNAME
        return Path.home() / f".{_APP_DIRNAME}"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / _APP_DIRNAME
    data_home = os.environ.get("XDG_DATA_HOME", str(Path.home() / ".local" / "share"))
    return Path(data_home) / _APP_DIRNAME


def converter_data_dir() -> Path:
    """转换器（vendored HTML→PPTX）的可变数据目录。

    环境变量 OFFIPY_CONVERTER_DATA_DIR 优先（测试注入/用户显式迁移），
    否则落在 user_data_dir()/converter 下，与 offipy 自身数据隔离。
    """
    override = os.environ.get(_CONVERTER_DATA_ENV)
    if override:
        return Path(override)
    return user_data_dir() / "converter"


def ensure_writable(path: str, overwrite: bool = False) -> str:
    """写入覆盖保护（P1 资源）：目标已存在且不显式 overwrite → FileExistsError。

    返回绝对路径（server 侧按调用方 CWD 无关）；save/save_pdf 的默认防线，
    防止 agent 或脚本无意间覆盖已有文件。
    """
    abs_path = os.path.abspath(path)
    if not overwrite and os.path.exists(abs_path):
        raise FileExistsError(f"目标文件已存在: {abs_path}（如确要覆盖请传 overwrite=True）")
    return abs_path
