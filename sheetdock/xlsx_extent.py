"""Stable data and display ranges, independent of cells allocated by rendering."""
from openpyxl.utils import column_index_from_string


def data_extent(sheet):
    rows, columns = 1, 1
    # Pinned openpyxl 3.1.5: reading via iter_rows would allocate blank cells.
    for cell in sheet._cells.values():
        if cell.value is not None or cell.has_style:
            rows, columns = max(rows, cell.row), max(columns, cell.column)
    for merged in sheet.merged_cells.ranges:
        rows, columns = max(rows, merged.max_row), max(columns, merged.max_col)
    return rows, columns


def extent(sheet):
    rows, columns = data_extent(sheet)
    for index in sheet.row_dimensions:
        rows = max(rows, index)
    for name, dimension in sheet.column_dimensions.items():
        columns = max(columns, column_index_from_string(name), dimension.min or 1, dimension.max or 1)
    return rows, columns


def extents(book):
    return tuple((sheet.title, extent(sheet)) for sheet in book)


def check_area(book):
    if sum(rows * columns for _, (rows, columns) in extents(book)) > 100_000:
        raise ValueError('表示範囲を含めたブック全体が10万セルを超えます')
