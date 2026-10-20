from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from openpyxl import Workbook

from sheetdock.csv_io import CsvDocument, CsvFormat
from sheetdock.gui import MultiFileApp, Session
from sheetdock.settings import SettingsStore
from sheetdock.tabs import Documents, path_key
from sheetdock.xlsx_gui import XlsxPane, XlsxSession
from sheetdock.xlsx_widths import MAX_EXCEL_WIDTH, width_to_pixels


@pytest.fixture
def csv_app(tmp_path):
    app = MultiFileApp.__new__(MultiFileApp)
    app.active = Documents().add(tmp_path / 'a.csv', Session(
        CsvDocument([['head', 'two'], ['a', 'b']], CsvFormat('utf-8', False, '\n', ',')), True))
    app.session, app.root = app.active.session, None
    app.sheet = SimpleNamespace(get_currently_selected=lambda: SimpleNamespace(row=0, column=1),
        get_column_widths=lambda: [120, 120], ops=SimpleNamespace(min_column_width=5))
    app.settings_store = SettingsStore(tmp_path / 'settings.json')
    app.render = Mock()
    return app


def test_csv_numeric_width_uses_selected_column_and_persists_clean_history(csv_app, monkeypatch):
    ask = Mock(return_value=230)
    monkeypatch.setattr('sheetdock.dialogs.askinteger', ask)
    csv_app.set_column_width()
    editor = csv_app.session.editor
    assert editor.column_widths == [120, 230] and not editor.dirty
    assert csv_app.settings_store.load().file_widths[path_key(csv_app.active.source)] == [120, 230]
    assert ask.call_args.kwargs['initialvalue'] == 120
    assert ask.call_args.kwargs['maxvalue'] == 5000
    editor.undo()
    assert editor.column_widths == [120, 120] and not editor.dirty
    editor.redo()
    assert editor.column_widths == [120, 230]
    csv_app.render.assert_called_once()


@pytest.mark.parametrize('value', [None, 120])
def test_csv_cancel_or_same_width_keeps_history_and_settings(csv_app, monkeypatch, value):
    monkeypatch.setattr('sheetdock.dialogs.askinteger', lambda *a, **kw: value)
    csv_app.set_column_width()
    assert not csv_app.session.editor._undo
    assert not csv_app.settings_store.path.exists()
    csv_app.render.assert_not_called()


def test_csv_width_requires_selection(csv_app, monkeypatch):
    csv_app.sheet.get_currently_selected = lambda: None
    ask, info = Mock(), Mock()
    monkeypatch.setattr('sheetdock.dialogs.askinteger', ask)
    monkeypatch.setattr('sheetdock.dialogs.showinfo', info)
    csv_app.set_column_width()
    ask.assert_not_called()
    info.assert_called_once()


@pytest.fixture
def pane(tmp_path):
    book = Workbook()
    book.active.append(['a', 'b'])
    book.active.column_dimensions['B'].width = 23.123456
    path = tmp_path / 'a.xlsx'
    book.save(path)
    pane = XlsxPane.__new__(XlsxPane)
    pane.session, pane.root, pane._digit_width = XlsxSession(path), None, 9
    pane.selected_cell = lambda: pane.session.sheet['B1']
    pane.render = Mock()
    return pane


def test_xlsx_numeric_width_roundtrips_through_book_and_history(pane, monkeypatch, tmp_path):
    before = pane.session.sheet.column_dimensions['B'].width
    ask = Mock(return_value=250)
    monkeypatch.setattr('sheetdock.dialogs.askinteger', ask)
    pane.set_column_width()
    assert pane.session.column_widths(9)[1] == 250 and pane.session.dirty
    assert ask.call_args.kwargs['maxvalue'] == width_to_pixels(MAX_EXCEL_WIDTH, 9)
    pane.session.history()
    assert pane.session.sheet.column_dimensions['B'].width == before and not pane.session.dirty
    pane.session.history(True)
    path = tmp_path / 'saved.xlsx'
    pane.session.save_as(path)
    assert XlsxSession(path).column_widths(9)[1] == 250
    pane.render.assert_called_once()


@pytest.mark.parametrize('cancel', [True, False])
def test_xlsx_cancel_or_same_width_keeps_original_precision(pane, monkeypatch, cancel):
    value = None if cancel else pane.session.column_widths(9)[1]
    monkeypatch.setattr('sheetdock.dialogs.askinteger', lambda *a, **kw: value)
    pane.set_column_width()
    assert pane.session.sheet.column_dimensions['B'].width == 23.123456
    assert not pane.session.dirty and not pane.session.undo_stack
    pane.render.assert_not_called()


@pytest.mark.parametrize('blocked', ['readonly', 'unselected'])
def test_xlsx_width_button_disabled_when_unavailable(pane, monkeypatch, blocked):
    if blocked == 'readonly':
        pane.session.probe.reasons.append('unsupported')
    else:
        pane.selected_cell = lambda: None
    pane.width_button, pane.height_button = Mock(), Mock()
    ask = Mock()
    monkeypatch.setattr('sheetdock.dialogs.askinteger', ask)
    pane.update_axis_commands()
    pane.width_button.configure.assert_called_once_with(state='disabled')
    pane.set_column_width()
    ask.assert_not_called()


def test_numeric_width_and_height_share_ordered_history(pane, monkeypatch):
    monkeypatch.setattr('sheetdock.dialogs.askinteger', lambda *a, **kw: 250)
    pane.session.resize_rows({0: 80})
    pane.set_column_width()
    pane.session.history()
    assert pane.session.row_heights()[0] == 80
    pane.session.history()
    assert not pane.session.dirty
