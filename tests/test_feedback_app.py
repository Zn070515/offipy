"""FeedbackApp：train/status 可用、has_com_root=False、顶层 numpy-free（F2-F）。"""

import subprocess
import sys

from offipy.feedback.app import FeedbackApp


def test_has_no_com_root():
    assert FeedbackApp().has_com_root is False


def test_status_returns_dict(tmp_path):
    res = FeedbackApp().status(feedback_dir=str(tmp_path))
    assert res["samples"] == 0
    assert res["model"] == "none"


def test_train_insufficient_pairs_returns_status(tmp_path):
    res = FeedbackApp().train(feedback_dir=str(tmp_path))
    assert res["trained"] is False


def test_feedback_app_lazy_import_no_numpy():
    """F2-F：从深路径 import FeedbackApp 也不拖 numpy（app.py 顶层仅标准库）。"""
    code = (
        "import sys\n"
        "from offipy.feedback.app import FeedbackApp\n"
        "assert 'numpy' not in sys.modules, 'app.py 顶层不得 import numpy'\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)


def test_import_offipy_no_numpy():
    code = (
        "import sys\n"
        "import offipy\n"
        "assert 'numpy' not in sys.modules, 'import offipy 不得拖 numpy'\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)
