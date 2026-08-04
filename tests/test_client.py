from offipy.client import convert_value


def test_convert_value_bool():
    assert convert_value("true") is True
    assert convert_value("false") is False


def test_convert_value_number():
    assert convert_value("42") == 42
    assert convert_value("3.14") == 3.14


def test_convert_value_none():
    assert convert_value("none") is None


def test_convert_value_str():
    assert convert_value("hello") == "hello"
