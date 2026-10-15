from types import SimpleNamespace

import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat
from sheetdock.gui import MultiFileApp, Session
from sheetdock.large_csv import FILE_WARNING_BYTES, CELL_WARNING_COUNT, large_file, large_table
from sheetdock.tabs import Documents


@pytest.mark.parametrize('size,expected', [(FILE_WARNING_BYTES - 1, False), (FILE_WARNING_BYTES, True), (FILE_WARNING_BYTES + 1, True)])
def test_file_warning_boundary(tmp_path, size, expected):
    path = tmp_path / 'large.csv'
    with path.open('wb') as stream:
        stream.truncate(size)
    assert large_file(path) == expected


@pytest.mark.parametrize('count,expected', [(CELL_WARNING_COUNT - 1, False), (CELL_WARNING_COUNT, True)])
def test_table_warning_boundary(count, expected):
    assert large_table([[''] * count]) == expected


def make_app(tmp_path):
    app = MultiFileApp.__new__(MultiFileApp)
    app.root = None
    app.documents = Documents()
    doc = CsvDocument([['ID'], ['00123']], CsvFormat('utf-8', False, '\n', ','))
    tab = app.documents.add(tmp_path / 'current.csv', Session(doc, True))
    tab.session.editor.set_cell(1, 0, '未保存')
    app.active = tab
    app.session = tab.session
    return app, tab


def test_file_warning_cancel_does_not_read_or_change_tabs(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = make_app(tmp_path)
    path = tmp_path / 'large.csv'
    with path.open('wb') as stream:
        stream.truncate(FILE_WARNING_BYTES)
    def unexpected(*a, **k):
        raise AssertionError('警告取消後に読込してはいけない')
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: False)
    monkeypatch.setattr(gui, 'load_checked', unexpected)
    app.open_path(str(path))
    assert app.documents.tabs == [tab]
    assert app.active is tab
    assert tab.session.editor.rows[1][0] == '未保存'
    assert tab.session.editor.dirty


def test_shape_warning_cancel_does_not_create_new_tab(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = make_app(tmp_path)
    path = tmp_path / 'many.csv'
    path.write_bytes(b'data')
    doc = CsvDocument([['x'] * CELL_WARNING_COUNT], CsvFormat('utf-8', False, '\n', ','))
    app.delimiter = SimpleNamespace(get=lambda: 'カンマ')
    app.requested_encoding = lambda: 'utf-8'
    monkeypatch.setattr(gui, 'load_checked', lambda *a, **k: (doc, None))
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: False)
    app.open_path(str(path))
    assert app.documents.tabs == [tab]
    assert app.session is tab.session
    assert tab.session.editor.dirty


def test_large_file_approval_skips_second_shape_prompt(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, _ = make_app(tmp_path)
    path = tmp_path / 'large.csv'
    with path.open('wb') as stream:
        stream.truncate(FILE_WARNING_BYTES)
    calls = []
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: calls.append(a) or True)
    assert app.confirm_large_file(path)
    doc = CsvDocument([['x'] * CELL_WARNING_COUNT], CsvFormat('utf-8', False, '\n', ','))
    assert app.confirm_large_table(path, doc)
    assert len(calls) == 1


def test_small_file_and_table_do_not_warn(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = make_app(tmp_path)
    path = tmp_path / 'small.csv'
    path.write_bytes(b'ID\n00123\n')
    def unexpected(*a, **k):
        raise AssertionError('小さいCSVには警告不要')
    monkeypatch.setattr(gui.messagebox, 'askyesno', unexpected)
    assert app.confirm_large_file(path)
    assert app.confirm_large_table(path, tab.session.snapshot())
