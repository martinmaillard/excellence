
from collections import UserList

import xlsxwriter

from .datastructures import DefaultList
from .exceptions import ImmutableCellError, InvalidGrid


class Workbook(xlsxwriter.Workbook):
    """Wrapper around :class:`xlsxwriter.Workbook` that provides to worksheets
    a reference to their parent workbook.
    """
    def add_worksheet(self, name=None):
        worksheet = super(Workbook, self).add_worksheet(name)

        # Add reference to workbook on worksheet
        worksheet._workbook = self

        return worksheet


class Grid():
    """Collection of rows that can be rendered in a :class:`Worksheet
    <xlsxwriter.Worksheet>`
    """
    def __init__(self, rows=None):
        self.rows = rows or []

    def write(self, worksheet, styles=None):
        # TODO: styles

        self._resolve_merged_cells()
        for row_index, row in enumerate(self.rows):
            for cell_index, cell in enumerate(row.cells):
                if cell.colspan > 1 or cell.rowspan > 1:
                    worksheet.merge_range(
                        row_index, cell_index,
                        row_index + cell.rowspan - 1,
                        cell_index + cell.colspan - 1,
                        cell.content
                    )
                else:
                    worksheet.write(row_index, cell_index, cell.content)

    def _resolve_merged_cells(self):
        """Fill positions that are part of a merged cell with :class:`NullCell`
        to make the rendering job trivial.

        Example of the applied transformation:

        First row:
            ---------------------
            |   1   |   | 3 |   |
            --------| 2 |---| 4 |
                    |   |   |   |
                    -----   -----
                                                -------------------------
        Second row:                             | 1 | . | 2 | 3 | 4 | . |
            -----------------           ==>     |------------------------
            | 5 | 6 |   7   |                   | 5 | 6 | . | 7 | . |
            -----------------                   |--------------------
                                                | 8 | . | 9 |
        Third row:                              |------------
            -------------
            |   8   | 9 |
            -------------

        """
        new_rows_cells = DefaultList(RowCells)
        for row_index, row in enumerate(self.rows):
            new_cells = new_rows_cells[row_index]

            cell_index = 0
            for cell in row.cells:
                cell_index = new_cells.next_empty_cell(cell_index)
                new_cells[cell_index] = cell

                for span_cell_index in range(
                    cell_index,
                    cell_index + cell.colspan
                ):
                    for span_row_index in range(
                        row_index,
                        row_index + cell.rowspan
                    ):
                        if (
                            span_cell_index == cell_index
                            and span_row_index == row_index
                        ):
                            continue

                        try:
                            new_rows_cells[span_row_index][span_cell_index] = \
                                NullCell(cell)
                        except ImmutableCellError:
                            raise InvalidGrid(
                                'Invalid grid definition: `{}` does not have '
                                'is too large for the available space.'.format(
                                    cell
                                )
                            )

            row.cells = new_cells


class Row():
    """Excel file row. Collection of `Cell`."""
    def __init__(self, cells=None):
        self.cells = cells or []


class Cell():
    """Excel file cell"""
    def __init__(self, content=None, colspan=1, rowspan=1, style=None):
        self.content = content
        self.colspan = colspan
        self.rowspan = rowspan
        self.style_name = style

    def __repr__(self):
        return 'Cell({})'.format(self.content)


class Style():
    """Collection of style properties with a name. Can extend another style
    defined earlier.
    """
    def __init__(self, name, properties, extends=None):
        self.name = name
        self.properties = properties
        self.extends = extends


class NullCell(Cell):
    """Used as a placeholder for positions occupied by merged cells in order to
    to keep them empty.
    """
    def __init__(self, extended_cell, **kw):
        super(NullCell, self).__init__(**kw)

        self.extended_cell = extended_cell

    def __repr__(self):
        return 'NullCell({})'.format(self.extended_cell)


class RowCells(UserList):
    """Acts like an infinite list of empty cells (None). Only allows writing in
    empty cells.
    """
    def __getitem__(self, index):
        try:
            return self.data[index]
        except IndexError:
            return None

    def __setitem__(self, index, data):
        current_value = self[index]

        if current_value is not None:
            raise ImmutableCellError
        else:
            for i in range(len(self.data), index + 1):
                self.data.append(None)

            self.data[index] = data

    def __iter__(self):
        return self.data.__iter__()

    def next_empty_cell(self, from_index):
        while self[from_index] is not None:
            from_index = from_index + 1

        return from_index
