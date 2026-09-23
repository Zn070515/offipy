"""Frozen entrypoint for the vendored HTML-to-PPTX converter."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path

_SOURCE_CONVERTER_DIR = Path(__file__).resolve().parent / "_vendor" / "html_to_editable_pptx"


def _converter_dir() -> Path:
    """Locate converter data in source mode or PyInstaller's extraction root."""

    if getattr(sys, "frozen", False):
        meipass = vars(sys).get("_MEIPASS")
        if not isinstance(meipass, str):
            raise RuntimeError("PyInstaller extraction root is unavailable")
        return Path(meipass) / "offipy" / "_vendor" / "html_to_editable_pptx"
    return _CONVERTER_DIR


_CONVERTER_DIR = _SOURCE_CONVERTER_DIR


def main() -> None:
    """Run the vendored converter as a script so its local imports keep working."""

    converter_dir = _converter_dir()
    converter_dir_text = str(converter_dir)
    if converter_dir_text not in sys.path:
        sys.path.insert(0, converter_dir_text)
    runpy.run_path(str(converter_dir / "convert.py"), run_name="__main__")


if __name__ == "__main__":
    main()
