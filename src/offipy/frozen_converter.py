"""Frozen entrypoint for the vendored HTML-to-PPTX converter."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path

_CONVERTER_DIR = Path(__file__).resolve().parent / "_vendor" / "html_to_editable_pptx"


def main() -> None:
    """Run the vendored converter as a script so its local imports keep working."""

    converter_dir = str(_CONVERTER_DIR)
    if converter_dir not in sys.path:
        sys.path.insert(0, converter_dir)
    runpy.run_path(str(_CONVERTER_DIR / "convert.py"), run_name="__main__")


if __name__ == "__main__":
    main()

