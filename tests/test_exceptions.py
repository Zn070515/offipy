"""异常体系测试：继承关系 + 统一可捕获（OffipyError 基类）+ 策略 A 领域异常。"""

import pytest

from offipy import (
    ComOperationError,
    ConversionError,
    FileConflictError,
    InvalidArgumentError,
    OfficeUnavailableError,
    OffipyError,
    ProtocolError,
    RemoteCallError,
    ServerStartError,
    TargetNotFoundError,
    UnsupportedPlatformError,
)

_ALL = [
    OfficeUnavailableError,
    ServerStartError,
    RemoteCallError,
    TargetNotFoundError,
    ConversionError,
    UnsupportedPlatformError,
    InvalidArgumentError,
    FileConflictError,
    ComOperationError,
    ProtocolError,
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


# --- 策略 A：领域异常带 code，RPC error_code 与异常一一对应（P1-4） ---


@pytest.mark.parametrize(
    ("exc", "code"),
    [
        (InvalidArgumentError, "invalid_argument"),
        (TargetNotFoundError, "target_not_found"),
        (FileConflictError, "file_conflict"),
        (ComOperationError, "com_operation"),
        (ProtocolError, "protocol"),
        (RemoteCallError, "remote_call"),
        (ServerStartError, "server_start"),
        (ConversionError, "conversion"),
        (OfficeUnavailableError, "office_unavailable"),
        (UnsupportedPlatformError, "unsupported_platform"),
    ],
)
def test_error_code_attribute(exc, code):
    assert exc.code == code
    assert getattr(exc("boom"), "code", None) == code


def test_invalid_argument_is_value_error():
    # 兼容既有 `except ValueError` 调用方：InvalidArgumentError 也是 ValueError
    assert issubclass(InvalidArgumentError, ValueError)
    with pytest.raises(ValueError):
        raise InvalidArgumentError("bad")


def test_file_conflict_is_file_exists_error():
    # 兼容既有 `except FileExistsError` 调用方
    assert issubclass(FileConflictError, FileExistsError)
    with pytest.raises(FileExistsError):
        raise FileConflictError("exists")


def test_com_operation_carries_hresult():
    e = ComOperationError("com fail", hresult=0x80010111, cause=ValueError("inner"))
    assert e.code == "com_operation"
    assert e.hresult == 0x80010111
    assert isinstance(e.cause, ValueError)
    # 缺省参数可空
    e2 = ComOperationError("plain")
    assert e2.hresult is None and e2.cause is None
