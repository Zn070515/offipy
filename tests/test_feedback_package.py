"""v1 feedback 包化：14 名透传 + numpy-free 顶层。"""

import subprocess
import sys

import offipy.feedback as feedback


def test_v1_names_reexported():
    for name in (
        "ALL_DIMENSIONS",
        "CONSISTENCY",
        "CONTRAST",
        "DEFAULT_DIR",
        "FEEDBACK_FILE",
        "PALETTE",
        "TYPE_SCALE",
        "VALID_ACTIONS",
        "WHITESPACE",
        "FeedbackRecord",
        "append",
        "dimension_weights",
        "load_records",
        "record_file",
    ):
        assert hasattr(feedback, name), name


def test_v1_private_constants_preserved():
    # 现有 test_feedback.py 访问这两个私有常量，必须保留
    assert feedback._WEIGHT_MAX == 3.0
    assert feedback._WEIGHT_MIN == 0.5


def test_import_offipy_feedback_no_numpy():
    """__init__ 顶层不拖 numpy（numpy-free 红线）。"""
    code = (
        "import sys\n"
        "import offipy.feedback\n"
        "assert 'numpy' not in sys.modules, 'feedback/__init__.py 顶层不得 import numpy'\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)
