from copy import deepcopy
from types import SimpleNamespace

import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat, load_csv
from sheetdock.gui import Session
from sheetdock.csv_tools_gui import CsvToolsApp, history_shortcut
from sheetdock.storage import fingerprint, save_checked
from sheetdock.tabs import Documents


ROWS = [['ID', '数量', '区分'], ['00123', '10', '東京'], ['00002', '2', '大阪'],
        ['00003', '2', '東京'], ['00004', '', '京都'], ['00005', '-1', '東京']]


def session(rows=None, header=True):
    return Session(CsvDocument(deepcopy(ROWS if rows is None else rows),
                               CsvFormat('utf-8', True, '\r\n', ',')), header)


def test_widget_retains_all_rows_while_filter_is_only_view_indices():
    s = session()
    s.editor.set_filter(2, query='東京')
    assert s.visible_body_indices() == [0, 2, 4]
    assert s.view() == (ROWS[0], ROWS[1:])
    assert s.snapshot().rows == ROWS and not s.editor.dirty
    s.editor.clear_filter()
    assert s.visible_body_indices() == [0, 1, 2, 3, 4]
    assert not s.editor.undo()


@pytest.mark.parametrize('header', [True, False])
def test_accept_full_widget_edit_and_condition_change(header):
    s = session(header=header)
    s.editor.set_filter(2, query='東京')
    _, body = s.view()
    body[3 - s.offset][2] = '北海道'
    assert s.accept_body(body)
    assert s.editor.rows[3][2] == '北海道' and s.editor.rows[2] == ROWS[2]
    assert s.visible_body_indices() == [1 - s.offset, 5 - s.offset]
    assert s.editor.undo() and s.editor.rows == ROWS
    assert s.editor.redo() and len(s.visible_body_indices()) == 2


@pytest.mark.parametrize('change', ['hidden_cell', 'row_removed', 'row_added'])
def test_hidden_rows_and_table_structure_cannot_change_through_widget(change):
    s = session()
    s.editor.set_filter(2, query='東京')
    body = s.view()[1]
    if change == 'hidden_cell':
        body[1][0] = 'wrong'
    elif change == 'row_removed':
        body.pop(1)
    else:
        body.append(['new'])
    with pytest.raises(ValueError):
        s.accept_body(body)
    assert s.editor.rows == ROWS and not s.editor.dirty and not s.editor.undo()


def test_visible_paste_changes_one_atomic_operation_even_when_targets_disappear():
    s = session()
    s.editor.set_filter(2, query='東京')
    assert s.editor.paste_visible(0, 1, [['101', '大阪'], ['202', '京都']])
    assert s.editor.rows[1][1:] == ['101', '大阪']
    assert s.editor.rows[3][1:] == ['202', '京都']
    assert s.editor.rows[2] == ROWS[2] and s.editor.rows[4] == ROWS[4]
    assert s.visible_body_indices() == [4]
    assert s.editor.undo() and not s.editor.dirty and not s.editor.undo()
    assert s.editor.redo() and s.visible_body_indices() == [4]


@pytest.mark.parametrize('row,col,values', [
    (2, 0, [['a'], ['b']]), (0, 2, [['a', 'b']]), (0, 0, [['a'], ['b', 'c']]),
    (0, 0, [[123]]), (-1, 0, [['a']]), (0, -1, [['a']]), (0, 0, [[]]),
])
def test_invalid_visible_paste_preserves_data_history_and_filter(row, col, values):
    s = session()
    s.editor.set_filter(2, query='東京')
    with pytest.raises((ValueError, IndexError, TypeError)):
        s.editor.paste_visible(row, col, values)
    assert s.editor.rows == ROWS and not s.editor.undo()
    assert s.editor.condition == (2, 'contains', '東京')


def test_zero_matches_still_saves_all_rows_and_savepoint(tmp_path):
    s = session()
    s.editor.set_cell(3, 0, '00099')
    s.editor.set_filter(2, query='不存在')
    assert s.visible_body_indices() == []
    path = tmp_path/'saved.csv'
    save_checked(s.snapshot(), path, fingerprint(path), backup=False)
    s.editor.mark_saved()
    restored = load_csv(path, encoding='utf-8')
    assert len(restored.rows) == 6 and restored.rows[3][0] == '00099'
    assert restored.format == s.format and s.editor.condition is not None
    assert s.editor.undo() and s.editor.dirty
    assert s.editor.redo() and not s.editor.dirty


def test_filter_state_is_per_tab_and_not_part_of_saved_data(tmp_path):
    documents = Documents()
    a = documents.add(tmp_path/'a.csv', session())
    b = documents.add(tmp_path/'b.csv', session())
    a.session.editor.set_filter(2, query='東京')
    assert b.session.editor.condition is None
    assert not a.session.editor.dirty and not b.session.editor.dirty
    assert a.session.snapshot() == b.session.snapshot()
    documents.saved_as(a, tmp_path/'copy.csv')
    assert a.session.editor.condition == (2, 'contains', '東京')


def test_header_toggle_requires_clearing_filter_then_syncs_sort_header():
    s = session()
    s.editor.set_filter(2, query='東京')
    with pytest.raises(ValueError, match='解除'):
        s.set_header_view(False)
    assert s.has_header and s.editor.header
    s.editor.clear_filter()
    s.set_header_view(False)
    assert not s.has_header and not s.editor.header
    s.editor.sort_rows(0)
    assert s.editor.rows[-1] == ROWS[0]


@pytest.mark.parametrize('operation,args', [('insert', (False,)), ('delete', (True,)),
    ('find', ()), ('search_action', ('all',)), ('toggle_header', ()),
    ('edit_header', ()), ('set_row_height', ())])
def test_filtered_unsafe_commands_stop_before_base_actions(monkeypatch, operation, args):
    notices = []
    app = CsvToolsApp.__new__(CsvToolsApp)
    s = session()
    s.editor.set_filter(2, query='東京')
    app.active = SimpleNamespace(session=s)
    app.root = None
    monkeypatch.setattr('sheetdock.csv_tools_gui.dialogs.showinfo', lambda *a, **k: notices.append(a))
    getattr(app, operation)(*args)
    assert notices and s.editor.rows == ROWS and not s.editor.dirty


@pytest.mark.parametrize('keysym,keycode,state,system,expected', [
    ('z', 52, 4, 'x11', False), ('Z', 52, 5, 'x11', True),
    ('??', 90, 4, 'win32', False), ('??', 90, 5, 'win32', True),
    ('z', 90, 0, 'win32', None), ('v', 86, 4, 'win32', None),
])
def test_production_history_dispatch_matches_validated_probe(keysym, keycode, state, system, expected):
    assert history_shortcut(SimpleNamespace(keysym=keysym, keycode=keycode, state=state), system) is expected


def test_reloaded_session_clears_inactive_widget_mapping_before_height_restore():
    s = session()
    s.editor.set_row_heights([None, 31, 42, 53, 64, 75])
    widget = SimpleNamespace(MT=SimpleNamespace(all_rows_displayed=False, min_row_height=10,
                                               get_default_row_height=lambda: 20))
    old_visible = [0, 2, 4]
    def display_rows(rows, **kwargs):
        assert rows == 'all'
        widget.MT.all_rows_displayed = True
    def heights(values):
        if not widget.MT.all_rows_displayed and len(values) != len(old_visible):
            raise IndexError('old native mapping cannot address restored row positions')
        widget.heights = values
    widget.display_rows = display_rows
    widget.set_row_heights = heights
    widget.row_index = lambda values: None
    widget.refresh = lambda: None
    app = CsvToolsApp.__new__(CsvToolsApp)
    app.apply_row_heights(SimpleNamespace(session=s), widget)
    assert widget.MT.all_rows_displayed and widget.heights == [31, 42, 53, 64, 75]
