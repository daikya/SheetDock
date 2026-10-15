from copy import deepcopy

import pytest

from sheetdock.editor import Editor
from sheetdock.csv_io import CsvDocument, CsvFormat, load_csv, save_csv


ROWS = [['名前', 'ID', 'メモ'], ['鈴木', '00123', '日本語\n日本語'], ['佐藤', '00000', ''], [], ['短い行']]


@pytest.mark.parametrize('operation', [
    lambda e: e.set_cell(1, 1, '00999'),
    lambda e: e.paste(1, 0, [['変更', '00001'], ['変更2', '00002']]),
    lambda e: e.insert_rows(1, 3),
    lambda e: e.delete_rows([1, 3, 1]),
    lambda e: e.insert_columns(2, 2),
    lambda e: e.delete_columns([0, 2]),
    lambda e: e.replace_all('日本語', '置換'),
])
def test_single_undo_restores_entire_action(operation):
    e = Editor(ROWS)
    operation(e)
    changed = e.rows
    assert changed != ROWS
    assert e.dirty
    assert e.undo()
    assert e.rows == ROWS
    assert not e.dirty
    assert not e.undo()
    assert e.redo()
    assert e.rows == changed
    assert not e.redo()


def test_history_sequence_and_redo_branch():
    e = Editor(ROWS)
    e.set_cell(1, 1, '00999')
    first = e.rows
    e.insert_rows(1, 2)
    second = e.rows
    e.replace_all('鈴木', '編集')
    assert e.undo() and e.rows == second
    assert e.undo() and e.rows == first
    assert e.redo() and e.rows == second
    e.set_cell(0, 0, '新しい編集')
    assert not e.redo()


def test_saved_state_tracks_undo_redo_by_content():
    e = Editor(ROWS)
    e.set_cell(1, 1, '00999')
    e.mark_saved()
    assert not e.dirty
    e.undo()
    assert e.dirty
    e.redo()
    assert not e.dirty


def test_copies_cannot_modify_history():
    source = deepcopy(ROWS)
    e = Editor(source)
    source[1][1] = '破壊'
    view = e.rows
    view[1][1] = '破壊'
    assert e.rows == ROWS
    payload = [['変更']]
    e.paste(1, 0, payload)
    payload[0][0] = '破壊'
    assert e.rows[1][0] == '変更'
    e.undo()
    assert e.rows == ROWS


@pytest.mark.parametrize('operation,error', [
    (lambda e: e.set_cell(1, 1, 123), TypeError),
    (lambda e: e.paste(4, 0, [['a'], ['b']]), IndexError),
    (lambda e: e.paste(0, 2, [['a', 'b']]), IndexError),
    (lambda e: e.delete_rows([1, 99]), IndexError),
    (lambda e: e.delete_columns([0, 99]), IndexError),
    (lambda e: e.insert_rows(-1, 2), IndexError),
    (lambda e: e.insert_columns(0, 0), ValueError),
    (lambda e: e.replace_all('', 'x'), ValueError),
    (lambda e: e.replace_all('鈴木', None), TypeError),
])
def test_invalid_operation_preserves_data_and_redo(operation, error):
    e = Editor(ROWS)
    e.set_cell(1, 0, '変更')
    e.undo()
    with pytest.raises(error):
        operation(e)
    assert e.rows == ROWS
    assert e.redo()
    assert e.rows[1][0] == '変更'


def test_noop_does_not_create_history_or_discard_redo():
    e = Editor(ROWS)
    e.set_cell(1, 0, '変更')
    e.undo()
    assert not e.set_cell(1, 0, '鈴木')
    assert not e.delete_rows([])
    assert e.replace_all('見つからない', 'x') == 0
    assert e.redo()


def test_find_and_replace_count_occurrences():
    e = Editor(ROWS)
    assert e.find('日本語') == [(1, 2)]
    assert e.replace_all('日本語', '置換') == 2
    assert e.rows[1][2] == '置換\n置換'
    assert e.rows[1][1] == '00123'


def test_delete_all_and_restore():
    e = Editor(ROWS)
    e.delete_rows(range(len(ROWS)))
    assert e.rows == []
    e.undo()
    assert e.rows == ROWS
    e.delete_columns(range(3))
    assert e.rows == [[] for _ in ROWS]
    e.undo()
    assert e.rows == ROWS


def test_empty_table_can_be_built():
    e = Editor([])
    e.insert_rows(0, 2)
    e.insert_columns(0, 2)
    e.set_cell(0, 0, '00123')
    assert e.rows == [['00123', ''], ['', '']]
    e.undo(); e.undo(); e.undo()
    assert e.rows == []


def test_edited_data_roundtrips_csv(tmp_path):
    e = Editor(ROWS)
    e.insert_columns(1, 2)
    e.replace_all('日本語', '編集')
    path = tmp_path / 'edited.csv'
    doc = CsvDocument(e.rows, CsvFormat('utf-8', True, '\r\n', ','))
    save_csv(doc, path)
    e.mark_saved()
    assert load_csv(path, encoding='utf-8') == doc
    assert not e.dirty
