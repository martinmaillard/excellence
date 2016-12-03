
from io import BytesIO

import pytest
import xlrd

from excellence import Cell, Grid, Row, Workbook, InvalidGrid


def write_grid(grid):
    output = BytesIO()

    with Workbook(output, {'in_memory': True}) as workbook:
        worksheet = workbook.add_worksheet()
        grid.write(worksheet)

    return output.getvalue()


def get_sheet(excel_content):
    wb = xlrd.open_workbook(file_contents=excel_content)
    return wb.sheet_by_index(0)


def test_empty_grid():
    grid = Grid()
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.nrows == 0


def test_one_cell_grid():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='Test'),
        ]),
    ])
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.cell_value(0, 0) == 'Test'


def test_multiple_rows_and_cells():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='A1'),
        ]),
        Row(cells=[
            Cell(content='B1'),
            Cell(content='B2'),
        ]),
    ])
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.cell_value(0, 0) == 'A1'
    assert sheet.cell_value(0, 0) == 'A1'
    assert sheet.cell_value(1, 0) == 'B1'
    assert sheet.cell_value(1, 1) == 'B2'


def test_cell_spanning_two_cols():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='Test', colspan=2),
        ]),
    ])
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.cell_value(0, 0) == 'Test'
    assert sheet.merged_cells == [(0, 1, 0, 2)]


def test_cell_spanning_two_rows():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='Test', rowspan=2),
        ]),
    ])
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.cell_value(0, 0) == 'Test'
    assert sheet.merged_cells == [(0, 2, 0, 1)]


def test_cell_spanning_multiple_rows_and_cols():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='Test', colspan=3, rowspan=4),
        ]),
    ])
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.cell_value(0, 0) == 'Test'
    assert sheet.merged_cells == [(0, 4, 0, 3)]


def test_spanning_cells_push_other_cells():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='A1', colspan=3, rowspan=4),
            Cell(content='A4'),
        ]),
        Row(cells=[
            Cell(content='B2'),
        ]),
        Row(cells=[
            Cell(content='C2'),
        ]),
    ])
    excel_content = write_grid(grid)

    sheet = get_sheet(excel_content)
    assert sheet.cell_value(0, 0) == 'A1'
    assert sheet.cell_value(0, 3) == 'A4'
    assert sheet.cell_value(1, 3) == 'B2'
    assert sheet.cell_value(2, 3) == 'C2'


def test_invalid_grid():
    grid = Grid(rows=[
        Row(cells=[
            Cell(content='A1', colspan=2),
            Cell(content='A3', rowspan=2),
        ]),
        Row(cells=[
            Cell(content='B1'),
            # There is only one space right besides B1 because A1 spans
            # two cols. This cell doesn't fit. The grid is invalid.
            Cell(content='B2', colspan=2),
        ]),
    ])

    with pytest.raises(InvalidGrid):
        write_grid(grid)


# TODO: styles
