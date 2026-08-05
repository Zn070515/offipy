"""PptApp set_title/set_body 空输入预校验（不碰 COM）。

用 __new__ 跳过 __init__（避免拉 Office），校验在触碰 COM 前抛
InvalidArgumentError——guard_com 只负责把 COM 失败归一为 ComOperationError。
"""

import pytest

from offipy.exceptions import InvalidArgumentError
from offipy.ppt import PptApp


def _app():
    return PptApp.__new__(PptApp)


def test_set_title_empty_rejected():
    app = _app()
    with pytest.raises(InvalidArgumentError):
        app.set_title(1, "")


def test_set_title_none_rejected():
    app = _app()
    with pytest.raises(InvalidArgumentError):
        app.set_title(1, None)


def test_set_body_empty_list_rejected():
    app = _app()
    with pytest.raises(InvalidArgumentError):
        app.set_body(1, [])


def test_set_body_none_rejected():
    app = _app()
    with pytest.raises(InvalidArgumentError):
        app.set_body(1, None)
