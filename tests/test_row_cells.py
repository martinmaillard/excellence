
import pytest

from excellence.core import RowCells
from excellence.exceptions import ImmutableCellError


def test_empty_row():
    cells = RowCells()
    assert cells == []


def test_can_read_any_empty_cell():
    cells = RowCells()

    assert cells[0] is None
    assert cells[10] is None
    assert cells[100] is None


def test_can_write_any_cell():
    cells = RowCells()

    cells[0] = 'test'
    assert cells[0] == 'test'

    cells[10] = 'test'
    assert cells[10] == 'test'


def test_cannot_write_in_filled_cell():
    cells = RowCells([1, 2, 3, 4, 5])

    with pytest.raises(ImmutableCellError):
        cells[3] = 'test'


def test_can_write_empty_cell_in_the_middle():
    cells = RowCells([1, 2, 3, None, 5])

    cells[3] = 'test'

    assert cells[3] == 'test'
    assert cells[4] == 5


def test_can_iterate():
    cells = RowCells([1, 2, 3])
    assert list(cells)


class TestNextEmptyCell():

    def test_empty_list(self):
        cells = RowCells()
        assert cells.next_empty_cell(0) == 0

    def test_list_with_values(self):
        cells = RowCells([1, 4, 3, 5])
        assert cells.next_empty_cell(0) == 4

    def test_empty_cell_in_the_middle(self):
        cells = RowCells([1, 4, 3, None, 5])
        assert cells.next_empty_cell(0) == 3

    def test_empty_cell_in_the_middle_from_the_middle(self):
        cells = RowCells([1, 4, 3, None, 5])
        assert cells.next_empty_cell(3) == 3

    def test_empty_cell_before_and_after(self):
        cells = RowCells([1, None, 3, None, 5])
        assert cells.next_empty_cell(3) == 3
