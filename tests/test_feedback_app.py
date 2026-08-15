"""FeedbackApp：train/status 可用、has_com_root=False、顶层 numpy-free（F2-F）。"""

import subprocess
import sys
from pathlib import Path

import pytest

from offipy.art.profiles import RULE_TITLE_TOO_SMALL
from offipy.audit import Severity
from offipy.exceptions import InvalidArgumentError
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


def test_train_requires_numpy_in_subprocess():
    code = (
        "import sys\n"
        "sys.modules['numpy'] = None\n"
        "from offipy.feedback.app import FeedbackApp\n"
        "try:\n"
        "    FeedbackApp().train()\n"
        "    raise SystemExit('should raise')\n"
        "except RuntimeError as e:\n"
        "    assert 'offipy[feedback]' in str(e)\n"
        "    print('ok')\n"
    )
    subprocess.run([sys.executable, "-c", code], check=True)


def test_train_rejects_non_dir(tmp_path):
    f = tmp_path / "not_a_dir"
    f.write_text("x")
    with pytest.raises(InvalidArgumentError):
        FeedbackApp().train(feedback_dir=str(f))


# ---------------------------------------------------------------- append


def test_append_writes_record(tmp_path):
    """FeedbackApp.append：severity 字符串 + features JSON 字符串 → 落记录。"""
    res = FeedbackApp().append(
        "balanced",
        RULE_TITLE_TOO_SMALL,
        "fixed",
        "MID",
        feedback_dir=str(tmp_path),
        features='{"finding.confidence": 0.5}',
    )
    assert Path(res["record"]).exists()
    from offipy.art.feedback import load_records

    recs = load_records(tmp_path)
    assert len(recs) == 1
    assert recs[0].rule_id == RULE_TITLE_TOO_SMALL
    assert recs[0].action == "fixed"
    assert recs[0].severity == Severity.MID
    assert recs[0].features == {"finding.confidence": 0.5}


def test_append_bad_features_json_raises(tmp_path):
    with pytest.raises(InvalidArgumentError, match="features"):
        FeedbackApp().append(
            "balanced",
            RULE_TITLE_TOO_SMALL,
            "fixed",
            "MID",
            feedback_dir=str(tmp_path),
            features="not-json",
        )
