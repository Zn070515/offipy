"""ppt.open_pres 纯逻辑（无 COM）：缺文件 pre-check → InvalidArgumentError。"""

import pytest

from offipy import ppt
from offipy.exceptions import InvalidArgumentError


def test_open_pres_missing_file_raises(tmp_path):
    # S5：open_pres 缺文件必须抛 InvalidArgumentError（对齐 word.open_doc 的 pre-check），
    # 不能等到 COM Presentations.Open 抛英文 ComOperationError。
    p = ppt.PptApp.__new__(ppt.PptApp)
    missing = tmp_path / "nope.pptx"
    with pytest.raises(InvalidArgumentError, match="源文件不存在"):
        p.open_pres(str(missing))
