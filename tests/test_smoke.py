"""
Simple pytest smoke tests to confirm the test runner is wired up.
"""

from pivot_util.pivot_util import _is_empty_value


def test_pytest_smoke() -> None:
    assert True


def test_is_empty_value() -> None:
    assert _is_empty_value(None) is True
    assert _is_empty_value("") is True
    assert _is_empty_value("   ") is True
    assert _is_empty_value(0) is False
    assert _is_empty_value(["", None, ["  "]]) is True
