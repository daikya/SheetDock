from copy import deepcopy

import pytest

from sheetdock.csv_tools import SortFilterEditor


ROWS = [['ID', '数量'], ['00123', '10'], ['00002', '2'],
        ['00003', '2'], ['00004', ''], ['00005', '-1']]


@pytest.mark.parametrize('mode,descending,ids', [
    ('text', False, ['00005', '00123', '00002', '00003', '00004']),
    ('text', True, ['00002', '00003', '00123', '00005', '00004']),
    ('number', False, ['00005', '00002', '00003', '00123', '00004']),
    ('number', True, ['00123', '00002', '00003', '00005', '00004']),
])
def test_stable_whole_row_sort_and_blank_last(mode, descending, ids):
    source = deepcopy(ROWS)
    editor = SortFilterEditor(source)
    assert editor.sort_rows(1, mode=mode, descending=descending)
    assert editor.rows[0] == ROWS[0]
    assert [row[0] for row in editor.rows[1:]] == ids
    assert sorted(editor.rows[1:]) == sorted(ROWS[1:])
    assert source == ROWS
    assert all(isinstance(value, str) for row in editor.rows for value in row)


def test_history_moves_heights_and_preserves_widths_and_savepoint():
    editor = SortFilterEditor(ROWS)
    editor.set_row_heights([20, 30, 40, 50, 60, 70])
    editor.set_column_widths([123, 234])
    editor.set_cell(1, 0, '変更')
    editor.mark_saved()
    saved = editor.rows
    editor.sort_rows(1, mode='number')
    assert editor.row_heights == [20, 70, 40, 50, 30, 60]
    sorted_rows = editor.rows
    assert editor.dirty and editor.column_widths == [123, 234]
    assert editor.undo() and editor.rows == saved and not editor.dirty
    assert editor.row_heights == [20, 30, 40, 50, 60, 70]
    assert editor.undo() and editor.rows == ROWS
    assert editor.redo() and editor.rows == saved
    assert editor.redo() and editor.rows == sorted_rows
    editor.mark_saved()
    assert not editor.dirty
    assert editor.undo() and editor.dirty


@pytest.mark.parametrize('value', ['abc', 'NaN', 'Infinity', '1,000', '１２', ' ', '1e999999999999999999999999'])
def test_bad_numeric_input_leaves_state_and_redo_untouched(value):
    editor = SortFilterEditor([['n'], ['2'], [value]])
    editor.set_cell(1, 0, '3')
    editor.undo()
    before = (editor.rows, editor.row_heights[:], editor.revision)
    with pytest.raises(ValueError):
        editor.sort_rows(0, mode='number')
    assert (editor.rows, editor.row_heights, editor.revision) == before
    assert editor.redo() and editor.rows[1] == ['3']


def test_decimal_precision_and_exponent_and_leading_zero():
    rows = [['0002'], ['2.0000000000000000000000000001'], ['+2e0'], ['.5']]
    editor = SortFilterEditor(rows, header=False)
    editor.sort_rows(0, mode='number')
    assert editor.rows == [rows[3], rows[0], rows[2], rows[1]]


@pytest.mark.parametrize('operator,query,indices', [
    ('contains', '2', [0, 2, 3]), ('equals', '2', [0, 2, 3]),
    ('empty', '', [0, 4]), ('nonempty', '', [0, 1, 2, 3, 5]),
    ('equals', 'none', [0]), ('contains', '', list(range(6))),
])
def test_filter_is_projection_only(operator, query, indices):
    editor = SortFilterEditor(ROWS)
    editor.set_cell(1, 0, '変更')
    editor.undo()
    revision = editor.revision
    editor.set_filter(1, operator=operator, query=query)
    assert editor.visible_indices() == indices
    assert editor.rows == ROWS and not editor.dirty and editor.revision == revision
    editor.clear_filter()
    assert editor.visible_indices() == list(range(6))
    assert editor.redo() and editor.rows[1][0] == '変更'


def test_headerless_ragged_blank_and_unicode_case_sensitive():
    editor = SortFilterEditor([['東京', 'B'], [], ['東京\n大阪', 'b'], ['別', 'A']], header=False)
    editor.set_filter(0, query='東京')
    assert editor.visible_indices() == [0, 2]
    editor.set_filter(1, operator='empty')
    assert editor.visible_indices() == [1]
    editor.set_filter(1, query='b')
    assert editor.visible_indices() == [2]
    editor.clear_filter()
    editor.sort_rows(1)
    assert editor.rows == [['別', 'A'], ['東京', 'B'], ['東京\n大阪', 'b'], []]


@pytest.mark.parametrize('rows,header', [([], True), ([['h']], True), ([['']], False)])
def test_empty_or_noop_has_no_history(rows, header):
    editor = SortFilterEditor(rows, header=header)
    if not rows:
        assert editor.visible_indices() == []
        with pytest.raises(IndexError):
            editor.sort_rows(0)
    else:
        assert not editor.sort_rows(0)
    assert not editor.undo() and not editor.dirty


@pytest.mark.parametrize('column', [-1, 2, True, 1.5])
def test_invalid_columns_are_atomic(column):
    editor = SortFilterEditor(ROWS)
    with pytest.raises(IndexError):
        editor.sort_rows(column)
    with pytest.raises(IndexError):
        editor.set_filter(column)
    assert not editor.undo() and editor.rows == ROWS
