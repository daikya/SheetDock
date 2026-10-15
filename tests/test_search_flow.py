from pathlib import Path
from types import SimpleNamespace

from sheetdock.csv_io import CsvDocument, CsvFormat
from sheetdock.gui import MultiFileApp, Session


def app():
    value = MultiFileApp.__new__(MultiFileApp)
    value.active = object()
    value.source = Path('data.csv')
    value.session = Session(CsvDocument([['a', 'ID'], ['aa', '00123']], CsvFormat('utf-8', False, '\n', ',')), True)
    value.snapshot = lambda: value.session.snapshot()
    value.render = lambda: None
    value.search_text = SimpleNamespace(get=lambda: 'a')
    value.replace_text = SimpleNamespace(get=lambda: 'b')
    value.messages = []
    value.search_result = SimpleNamespace(set=value.messages.append)
    value.selection = []
    value.sheet = SimpleNamespace(select_column=lambda c: value.selection.append(('header', c)), select_cell=lambda r, c: value.selection.append(('body', r, c)), see=lambda *a: None)
    return value


def test_next_previous_select_header_and_body():
    a = app()
    a.search_action('next')
    assert a.selection[-1] == ('header', 0)
    assert '1/3' in a.messages[-1]
    a.search_action('next')
    assert a.selection[-1] == ('body', 0, 0)
    assert '2/3' in a.messages[-1]
    a.search_action('previous')
    assert a.selection[-1] == ('header', 0)


def test_one_replacement_undo_restores_one_occurrence():
    a = app()
    a.search_action('next')
    a.search_action('next')
    a.search_action('one')
    assert a.session.editor.rows == [['a', 'ID'], ['ba', '00123']]
    assert a.session.search.current.start == 1
    assert a.session.editor.undo()
    assert a.session.editor.rows == [['a', 'ID'], ['aa', '00123']]


def test_all_replacement_includes_header_and_is_one_action():
    a = app()
    a.search_action('all')
    assert a.session.editor.rows == [['b', 'ID'], ['bb', '00123']]
    assert '3箇所' in a.messages[-1]
    assert a.session.editor.undo()
    assert a.session.editor.rows == [['a', 'ID'], ['aa', '00123']]


def test_empty_search_reports_error_without_edit():
    a = app()
    a.search_text = SimpleNamespace(get=lambda: '')
    a.search_action('one')
    assert '検索文字列' in a.messages[-1]
    assert not a.session.editor.dirty
