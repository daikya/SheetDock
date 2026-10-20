from types import SimpleNamespace

import pytest

from sheetdock.editor import Editor
from sheetdock.gui import MultiFileApp, Session
from sheetdock.csv_io import CsvDocument, CsvFormat
from sheetdock.settings import SettingsStore
from sheetdock.tabs import Documents, path_key
from sheetdock.widths import fit_widths


def test_width_resize_changes_neither_data_nor_dirty_history():
    e = Editor([['a', 'b', 'c']])
    e.set_column_widths([80, 240, 360])
    assert e.column_widths == [80, 240, 360]
    assert e.rows == [['a', 'b', 'c']]
    assert not e.dirty
    assert not e.undo()


def test_insert_columns_moves_widths_and_undo_redo_restores():
    e = Editor([['a', 'b', 'c']])
    e.set_column_widths([80, 240, 360])
    e.insert_columns(1, 2)
    assert e.column_widths == [80, 120, 120, 240, 360]
    e.undo()
    assert e.column_widths == [80, 240, 360]
    e.redo()
    assert e.column_widths == [80, 120, 120, 240, 360]


def test_delete_columns_and_undo_redo_restore_widths():
    e = Editor([['a', 'b', 'c']])
    e.set_column_widths([80, 240, 360])
    e.delete_columns([0, 2])
    assert e.column_widths == [240]
    e.undo()
    assert e.column_widths == [80, 240, 360]
    e.redo()
    assert e.column_widths == [240]


def test_cell_edit_undo_keeps_later_manual_resize():
    e = Editor([['a', 'b']])
    e.set_cell(0, 0, '変更')
    e.set_column_widths([200, 400])
    e.undo()
    assert e.column_widths == [200, 400]
    e.redo()
    assert e.column_widths == [200, 400]


def test_all_rows_deleted_and_undo_restore_widths():
    e = Editor([['a', 'b']])
    e.set_column_widths([200, 400])
    e.delete_rows([0])
    assert e.column_widths == []
    e.undo()
    assert e.column_widths == [200, 400]


@pytest.mark.parametrize('widths,count,expected', [([80, 200], 3, [80, 200, 120]), ([80, 200, 300], 1, [80]), ([True, -1, float('nan'), 9000, '80'], 5, [120] * 5), (None, 2, [120, 120])])
def test_invalid_or_changed_column_counts(widths, count, expected):
    assert fit_widths(widths, count) == expected


def test_file_widths_are_independent_and_survive_other_updates(tmp_path):
    store = SettingsStore(tmp_path / 'settings.json')
    a, b = tmp_path / 'a.csv', tmp_path / 'b.csv'
    store.update(widths_file=a, column_widths=[80, 300])
    store.update(widths_file=b, column_widths=[240, 100])
    store.update(backup_enabled=False, window_geometry='900x600+10+20')
    loaded = store.load()
    assert loaded.file_widths[path_key(a)] == [80, 300]
    assert loaded.file_widths[path_key(b)] == [240, 100]
    assert not loaded.backup_enabled


def test_widths_restored_to_reloaded_column_count(tmp_path):
    store = SettingsStore(tmp_path / 'settings.json')
    path = tmp_path / 'a.csv'
    store.update(widths_file=path, column_widths=[80, 300, 400])
    tab = Documents().add(path, Session(CsvDocument([['a', 'b']], CsvFormat('utf-8', False, '\n', ',')), True))
    app = MultiFileApp.__new__(MultiFileApp)
    app.settings_store = store
    app.load_widths(tab)
    assert tab.session.editor.column_widths == [80, 300]


def test_capture_widths_and_dirty_structure_not_written_until_save(tmp_path):
    app = MultiFileApp.__new__(MultiFileApp)
    app.settings_store = SettingsStore(tmp_path / 'settings.json')
    path = tmp_path / 'a.csv'
    tab = Documents().add(path, Session(CsvDocument([['a', 'b']], CsvFormat('utf-8', False, '\n', ',')), True))
    app.pages = {tab: (None, SimpleNamespace(get_column_widths=lambda: [80, 300]))}
    app.applying_widths = False
    app.capture_widths(tab, persist=True)
    assert app.settings_store.load().file_widths[path_key(path)] == [80, 300]
    tab.session.editor.insert_columns(1, 1)
    app.persist_widths(tab)
    assert app.settings_store.load().file_widths[path_key(path)] == [80, 300]
    tab.session.editor.mark_saved()
    app.persist_widths(tab)
    assert app.settings_store.load().file_widths[path_key(path)] == [80, 120, 300]


def test_header_toggle_preserves_widths():
    s = Session(CsvDocument([['a', 'b'], ['1', '2']], CsvFormat('utf-8', False, '\n', ',')), True)
    s.editor.set_column_widths([80, 300])
    s.set_header_view(False)
    assert s.editor.column_widths == [80, 300]
    s.set_header_view(True)
    assert s.editor.column_widths == [80, 300]


def test_width_history_does_not_dirty_csv_and_redoes():
    editor = Editor([['a', 'b']])
    assert editor.resize_columns([240, 180])
    assert not editor.dirty
    assert editor.undo()
    assert editor.column_widths == [120, 120]
    assert not editor.dirty
    assert editor.redo()
    assert editor.column_widths == [240, 180]
    assert not editor.dirty


def test_width_cell_and_structure_history_follow_operation_order():
    editor = Editor([['a', 'b']])
    editor.resize_columns([240, 180])
    editor.set_cell(0, 0, 'edited')
    editor.insert_columns(1, 1)
    editor.resize_columns([260, 150, 200])
    editor.undo()
    assert editor.column_widths == [240, 120, 180]
    editor.undo()
    assert editor.column_widths == [240, 180]
    editor.undo()
    assert editor.rows == [['a', 'b']] and not editor.dirty
    assert editor.column_widths == [240, 180]
    editor.undo()
    assert editor.column_widths == [120, 120]
    for _ in range(4):
        assert editor.redo()
    assert editor.rows == [['edited', '', 'b']]
    assert editor.column_widths == [260, 150, 200]


def test_width_noop_keeps_redo_but_new_resize_clears_it():
    editor = Editor([['a', 'b']])
    editor.resize_columns([240, 180])
    editor.undo()
    assert not editor.resize_columns([120, 120])
    assert editor.redo()
    editor.undo()
    editor.resize_columns([260, 190])
    assert not editor.redo()


def test_csv_drag_records_final_width_once_and_persists_undo(tmp_path):
    app = MultiFileApp.__new__(MultiFileApp)
    app.settings_store = SettingsStore(tmp_path / 'settings.json')
    tab = Documents().add(tmp_path / 'a.csv', Session(
        CsvDocument([['a', 'b']], CsvFormat('utf-8', False, '\n', ',')), False))
    header = SimpleNamespace(currently_resizing_width=True)
    widths = [120, 120]
    sheet = SimpleNamespace(CH=header, get_column_widths=lambda: widths[:])
    app.pages = {tab: (None, sheet)}
    app.rendering = app.applying_widths = False
    for pixels in (140, 160, 180, 200):
        widths[0] = pixels
        app.page_resized(tab)
    assert not tab.session.editor._undo
    header.currently_resizing_width = False
    app.page_resized(tab)
    assert len(tab.session.editor._undo) == 1
    assert not tab.session.editor.dirty
    assert app.settings_store.load().file_widths[path_key(tab.source)] == [200, 120]
    tab.session.editor.undo()
    app.persist_widths(tab)
    assert app.settings_store.load().file_widths[path_key(tab.source)] == [120, 120]
