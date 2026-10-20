import csv
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from sheetdock.xlsx_gui import XlsxPane, XlsxSession
from sheetdock.xlsx_inline import InlineBridge
from sheetdock.xlsx_io import snapshot, ProtectedCellError
from sheetdock.mixed_gui import SheetDockApp
from xlsx_sort_fixtures import make_samples


@pytest.fixture
def pane(tmp_path):
    make_samples(tmp_path)
    pane = XlsxPane.__new__(XlsxPane)
    pane.session = XlsxSession(tmp_path/'sort_basic.xlsx')
    pane.root = None
    pane.session.set_filter(5, query='東京')
    pane.selection = SimpleNamespace(row=2, column=0)
    pane.indices = pane.session.displayed_indices()
    pane.sheet_widget = SimpleNamespace(
        get_currently_selected=lambda: pane.selection,
        display_rows=lambda: pane.indices)
    def render():
        pane.indices = pane.session.displayed_indices()
    pane.render = Mock(side_effect=render)
    yield pane
    pane.session.probe.book.close()


@pytest.mark.parametrize('row,address', [(0, 'A1'), (1, 'A2'), (2, 'A4'), (3, 'A6'), (4, None), (-1, None)])
def test_current_display_selection_maps_to_workbook_row(pane, row, address):
    pane.selection.row = row
    cell = pane.selected_cell()
    assert (cell.coordinate if cell else None) == address
    if cell and row == 0:
        assert pane.filtered_cell_reason(cell)


def test_inline_edit_maps_data_row_and_refilters_on_undo(pane):
    bridge = InlineBridge(pane.session, Mock(), Mock(), pane.inline_coordinates)
    event = {'row': 2, 'column': 4, 'key': 'F2', 'value': '東京'}
    assert bridge.begin(event) == '東京'
    assert bridge.validate({**event, 'value': '北海道'}) == '北海道'
    s = pane.session
    assert s.sheet['E4'].value == '北海道' and s.sheet['E3'].value == '大阪'
    assert s.visible_rows() == [2, 6] and len(s.undo_stack) == 1
    # The old native view cannot be used as an edit target until it is refreshed.
    assert pane.selected_cell() is None
    assert s.history() and s.visible_rows() == [2, 4, 6] and not s.dirty
    assert s.history(True) and s.visible_rows() == [2, 6]


@pytest.mark.parametrize('change', ['filter', 'sheet', 'book', 'cell'])
def test_pending_inline_input_rejects_changed_context(pane, change):
    s = pane.session
    bridge = InlineBridge(s, Mock(), Mock(), pane.inline_coordinates)
    event = {'row': 2, 'column': 0, 'key': 'F2', 'value': '00003'}
    assert bridge.begin(event) == '00003'
    if change == 'filter':
        s.set_filter(5, query='大阪')
    elif change == 'sheet':
        s.sheet_name = 'Other'
    elif change == 'book':
        s.clear_filter()
        s.sort_rows(2)
    else:
        s.edit(3, 0, 'changed')
    before, history = snapshot(s.probe.book), list(s.undo_stack)
    assert bridge.validate({**event, 'value': 'stale'}) is None
    assert snapshot(s.probe.book) == before and s.undo_stack == history


def test_filtered_header_cannot_be_directly_edited_or_pasted(pane):
    pane.selection.row = 0
    bridge = InlineBridge(pane.session, Mock(), Mock(), pane.inline_coordinates)
    assert bridge.begin({'row': 0, 'column': 0, 'key': 'F2', 'value': 'コード'}) is None
    for operation in [pane.session.edit, pane.session.paste]:
        with pytest.raises(ValueError):
            operation(0, 0, 'header')
    assert not pane.session.undo_stack and not pane.session.dirty


def test_filtered_paste_targets_stable_noncontiguous_rows_once(pane):
    s = pane.session
    assert s.paste(1, 4, '大阪\tfirst\n京都\tsecond')
    assert s.sheet['E2'].value == '大阪' and s.sheet['F2'].value == 'first'
    assert s.sheet['E4'].value == '京都' and s.sheet['F4'].value == 'second'
    assert s.sheet['E3'].value == '大阪' and s.sheet['E5'].value == '京都'
    assert s.visible_rows() == [6] and len(s.undo_stack) == 1
    assert s.history() and s.visible_rows() == [2, 4, 6] and not s.dirty
    assert s.history(True) and s.visible_rows() == [6]


@pytest.mark.parametrize('row,column,text', [(5, 0, 'a\nb'), (1, 5, 'a\tb'),
    (1, 0, 'new\tabc'), (1, 0, 'new\tNaN'), (1, 0, 'new\tInfinity'),
    (1, 4, 'new\t=1+1'), (1, 0, 'a\tb\nc'), (2, 0, 'hidden')])
def test_filtered_paste_failure_preserves_all_cells_and_redo(pane, row, column, text):
    s = pane.session
    s.edit(1, 0, 'changed')
    s.history()
    before, undo, redo = snapshot(s.probe.book), list(s.undo_stack), list(s.redo_stack)
    with pytest.raises((ValueError, TypeError, ProtectedCellError)):
        s.paste(row, column, text)
    assert snapshot(s.probe.book) == before and s.undo_stack == undo and s.redo_stack == redo


@pytest.mark.parametrize('operation', ['find', 'replace', 'structure', 'axis', 'height'])
def test_hidden_data_operations_require_filter_clear(pane, operation):
    s = pane.session
    before = snapshot(s.probe.book)
    actions = {'find': lambda: s.find('東京'), 'replace': lambda: s.replace_all('東京', 'new'),
               'structure': lambda: s.sheet_operation('add', 'New'),
               'axis': lambda: s.axis_operation('delete-rows', 2),
               'height': lambda: s.resize_rows({1: 60})}
    with pytest.raises((ValueError, ProtectedCellError), match='解除'):
        actions[operation]()
    assert snapshot(s.probe.book) == before and not s.undo_stack


def test_other_sheet_filter_blocks_whole_book_replace(pane):
    pane.session.sheet_name = 'Other'
    with pytest.raises(ValueError, match='解除'):
        pane.session.replace_all('他', 'new', all_sheets=True)
    assert pane.session.sheet['A1'].value == '他シート不変'


@pytest.mark.parametrize('query', ['東京', '該当なし'])
def test_filtered_save_and_export_include_all_rows_and_other_sheets(pane, tmp_path, query):
    s = pane.session
    source = s.probe.path
    raw = source.read_bytes()
    s.edit(1, 0, '00099')
    s.set_filter(5, operator='equals', query=query)
    filters = dict(s.filters)
    expected = snapshot(s.probe.book)
    target = tmp_path/'copy.xlsx'
    s.save_as(target)
    reopened = XlsxSession(target)
    try:
        assert snapshot(reopened.probe.book) == expected
        assert not reopened.filters and reopened.sheet['A2'].value == '00099'
        assert reopened.probe.book['Other']['A1'].value == '他シート不変'
        assert reopened.sheet.auto_filter.ref is None
    finally:
        reopened.probe.book.close()
    assert not s.dirty and s.filters == filters
    export = tmp_path/'all.csv'
    s.export_csv(export)
    with export.open(encoding='utf-8-sig', newline='') as stream:
        rows = list(csv.reader(stream))
    assert len(rows) == 6 and rows[2][0] == '00002' and rows[4][0] == '00004'
    assert s.filters == filters and not s.dirty and source.read_bytes() == raw


def test_save_failure_preserves_filter_path_stamp_dirty_and_history(pane, monkeypatch):
    s = pane.session
    s.edit(3, 0, 'changed')
    before = snapshot(s.probe.book), s.probe.path, s.stamp, dict(s.filters), list(s.undo_stack)
    monkeypatch.setattr('sheetdock.xlsx_gui.save_checked', Mock(side_effect=PermissionError('locked')))
    with pytest.raises(PermissionError):
        s.save()
    assert (snapshot(s.probe.book), s.probe.path, s.stamp, s.filters, s.undo_stack) == before
    assert s.dirty


def test_structure_undo_clears_memory_filters_to_avoid_name_reuse(pane):
    s = pane.session
    s.clear_filter()
    s.sheet_operation('add', 'New')
    s.set_filter(1, query='none', header=False)
    s.history()
    assert 'New' not in s.probe.book.sheetnames and not s.filters
    s.history(True)
    assert 'New' in s.probe.book.sheetnames and not s.filters


def test_sort_edit_width_height_history_remains_ordered_and_filter_is_not_history(pane):
    s = pane.session
    s.clear_filter()
    baseline = snapshot(s.probe.book)
    s.edit(1, 0, '00099')
    s.resize_rows({1: 60})
    s.resize_columns({0: 240})
    s.sort_rows(2)
    sorted_state = snapshot(s.probe.book)
    s.set_filter(5, query='東京')
    for _ in range(4):
        assert s.history()
    assert snapshot(s.probe.book) == baseline and s.visible_rows() == [2, 4, 6] and not s.dirty
    for _ in range(4):
        assert s.history(True)
    assert snapshot(s.probe.book) == sorted_state and s.dirty and len(s.undo_stack) == 4


def test_header_false_includes_first_row_and_keeps_choice_per_sheet(pane):
    s = pane.session
    s.clear_filter()
    s.set_filter(1, query='コード', header=False)
    assert s.displayed_indices() == [0] and s.visible_rows() == [1]
    assert s.filter_headers['Data'] is False
    s.edit(0, 0, '見出しも本文')
    assert s.visible_rows() == []
    s.sheet_name = 'Other'
    assert not s.filtered
    s.sheet_name = 'Data'
    assert s.filtered and s.visible_rows() == []


@pytest.mark.parametrize('system,keysym,keycode,state,redo', [
    ('win32', 'z', 90, 4, False), ('win32', 'Z', 90, 5, True),
    ('win32', 'unknown', 90, 4, False), ('win32', 'unknown', 90, 5, True),
    ('x11', 'z', 52, 4, False), ('x11', 'Z', 52, 5, True)])
def test_shared_first_key_tag_routes_xlsx_history_once(pane, system, keysym, keycode, state, redo):
    app = SheetDockApp.__new__(SheetDockApp)
    app.active = SimpleNamespace(session=pane.session)
    app.root = SimpleNamespace(tk=SimpleNamespace(call=lambda *a: system))
    app.history = Mock()
    event = SimpleNamespace(widget=SimpleNamespace(winfo_class=lambda: 'TButton'),
                            keysym=keysym, keycode=keycode, state=state)
    assert app.csv_key(event) == 'break'
    app.history.assert_called_once_with(redo)


@pytest.mark.parametrize('widget_class', ['Text', 'TEntry', 'TCombobox'])
def test_shared_history_key_leaves_xlsx_input_fields_alone(pane, widget_class):
    app = SheetDockApp.__new__(SheetDockApp)
    app.active = SimpleNamespace(session=pane.session)
    app.history = Mock()
    event = SimpleNamespace(widget=SimpleNamespace(winfo_class=lambda: widget_class),
                            keysym='z', keycode=90, state=4)
    assert app.csv_key(event) is None
    app.history.assert_not_called()


def test_axis_history_clears_conditions_with_obsolete_column_references(pane):
    s = pane.session
    s.clear_filter()
    s.append_axis('columns', 1)
    s.set_filter(7, operator='empty')
    assert s.filtered
    s.history()
    assert not s.filters and len(s.column_widths()) == 6
    s.history(True)
    assert not s.filters and len(s.column_widths()) == 7


def test_filtered_heading_does_not_redisplay_originally_hidden_first_row(pane):
    s = pane.session
    s.sheet.row_dimensions[1].hidden = True
    assert s.displayed_indices() == [1, 3, 5]
    assert s.sheet.row_dimensions[1].hidden
