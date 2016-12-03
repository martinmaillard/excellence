
import pytest

from excellence.datastructures import DefaultList


_empty = object()


def empty():
    return _empty


def test_no_default_factory_raises():
    l = DefaultList()

    with pytest.raises(IndexError):
        l[0]


def test_existing_index():
    l = DefaultList(empty, [1, 2, 3])
    assert l[1] == 2


def test_get_missing_index():
    l = DefaultList(empty)
    assert l[0] == _empty


def test_get_further_missing_index():
    l = DefaultList(empty)
    assert l[4] == _empty


def test_insert_missing_index():
    l = DefaultList(empty)
    l[0] = 34
    assert l[0] == 34


def test_insert_further_missing_index():
    l = DefaultList(empty)
    l[2] = 34
    assert l[2] == 34
    assert l[0] == _empty
