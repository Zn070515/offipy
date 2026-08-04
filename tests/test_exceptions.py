"""异常体系测试：继承关系 + 统一可捕获（OffipyError 基类）。"""

import pytest

from offipy import (
    ConversionError,
    OfficeUnavailableError,
    OffipyError,
    RemoteCallError,
    ServerStartError,
    UnsupportedPlatformError,
)

_ALL = [
    OfficeUnavailableError,
    ServerStartError,
    RemoteCallError,
    ConversionError,
    UnsupportedPlatformError,
]


def test_offipy_error_is_exception():
    assert issubclass(OffipyError, Exception)


@pytest.mark.parametrize("exc", _ALL)
def test_exception_subclass_of_offipy(exc):
    assert issubclass(exc, OffipyError)


@pytest.mark.parametrize("exc", _ALL)
def test_exception_catchable_as_base(exc):
    try:
        raise exc("boom")
    except OffipyError as e:
        assert str(e) == "boom"
