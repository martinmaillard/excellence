
# Excellence

Excellence is an attempt at a pleasant way to create Excel files with Python.
It's actually just a wrapper around
[XlsxWriter](https://github.com/jmcnamara/XlsxWriter) with a more declarative
API.

Example:

    from excellence import Grid, Style, Workbook

    grid = Grid(rows=[
        Row(cells=[
            Cell(content='My title', colspan=3, rowspan=4, style='title'),
        ]),
        Row(height=40, cells=[
            Cell(content='First col', style='header'),
            Cell(content='Second col', style='header'),
            Cell(content='Third col', style='header'),
        ]),
        Row(cells=[
            Cell(content='My value'),
            Cell(content='Another value', style='number'),
            Cell(content='And a third one'),
        ]),
    ])

    styles = [
        Style('header', {
            'bold': True,
        }),
        Style('title', {
            'font_size': 22,
        }, extends='header'),
        Style('number', {
            'num_format': '0.00',
        }),
    ]

    with Workbook(output, {'in_memory': True}) as workbook:
        worksheet = workbook.add_worksheet()
        grid.write(worksheet, styles=styles)
