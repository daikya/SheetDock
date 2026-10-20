from copy import deepcopy
from types import SimpleNamespace

import pytest

from probes.csv_sort_filter_gui_probe import ProjectionGui, ProjectionSession, history_shortcut
from probes.csv_sort_filter_probe import make_samples
from sheetdock.csv_io import CsvDocument, CsvFormat, load_csv


ROWS = [['ID', '数量', '区分'], ['00123', '10', '東京'], ['00002', '2', '大阪'],
        ['00003', '2', '東京'], ['00004', '', '京都'], ['00005', '-1', '東京']]


def session(rows=None, header=True):
    return ProjectionSession(CsvDocument(deepcopy(ROWS if rows is None else rows),
                                        CsvFormat('utf-8', True, '\r\n', ',')), header=header)


def test_filtered_edit_targets_data_row_and_reevaluates_on_history():
    s = session()
    s.editor.set_filter(2, query='東京')
    header, body, indices, token = s.view()
    assert header == ROWS[0] and indices == [1, 3, 5]
    assert body == [ROWS[i] for i in indices]
    assert s.edit(1, 0, '00099', token)
    assert s.editor.rows[3] == ['00099', '2', '東京']
    assert s.editor.rows[2] == ROWS[2] and s.editor.rows[4] == ROWS[4]
    _, _, _, token = s.view()
    assert s.edit(1, 2, '北海道', token)
    assert s.view()[2] == [1, 5]
    assert s.editor.undo() and s.view()[2] == [1, 3, 5]
    assert s.editor.undo() and s.editor.rows == ROWS and not s.editor.dirty
    assert s.editor.redo() and s.editor.redo() and s.view()[2] == [1, 5]


def test_paste_maps_all_rows_before_filter_changes_atomically():
    s = session()
    s.editor.set_filter(2, query='東京')
    token = s.view()[3]
    assert s.paste(0, 1, [['101', '大阪'], ['202', '京都']], token)
    assert s.editor.rows[1] == ['00123', '101', '大阪']
    assert s.editor.rows[3] == ['00003', '202', '京都']
    assert s.editor.rows[2] == ROWS[2] and s.editor.rows[4] == ROWS[4]
    assert s.view()[2] == [5]
    assert s.editor.undo() and s.editor.rows == ROWS
    assert s.editor.redo() and s.view()[2] == [5]


@pytest.mark.parametrize('change', ['filter', 'sort', 'edit', 'history'])
def test_stale_selection_rejected(change):
    s = session()
    token = s.view()[3]
    if change == 'filter':
        s.editor.set_filter(2, query='東京')
    elif change == 'sort':
        s.editor.sort_rows(1, mode='number')
    else:
        s.editor.set_cell(1, 0, 'new')
        if change == 'history':
            s.editor.undo()
    before = s.editor.rows
    with pytest.raises(ValueError, match='選び直し'):
        s.edit(0, 0, 'wrong', token)
    assert s.editor.rows == before


@pytest.mark.parametrize('row,column,values', [
    (3, 0, [['x']]), (2, 0, [['x'], ['y']]), (0, 2, [['x', 'y']]),
    (0, 0, [['x'], ['y', 'z']]), (0, 0, [[123]]), (0, 0, []),
    (0, 0, [[]]), (-1, 0, [['x']]), (0, -1, [['x']]),
])
def test_bad_paste_rejects_everything(row, column, values):
    s = session()
    s.editor.set_filter(2, query='東京')
    token = s.view()[3]
    with pytest.raises((IndexError, ValueError)):
        s.paste(row, column, values, token)
    assert s.editor.rows == ROWS and not s.editor.dirty and not s.editor.undo()


def test_zero_matches_and_header_only_never_become_editable_data():
    for rows in (ROWS, [ROWS[0]], []):
        s = session(rows)
        if rows:
            s.editor.set_filter(0, operator='equals', query='不存在')
        _, body, indices, token = s.view()
        assert body == [] and indices == []
        with pytest.raises(IndexError):
            s.edit(0, 0, 'wrong', token)
        assert s.editor.rows == rows


def test_headerless_ragged_edit_keeps_hidden_and_unedited_rows_exact():
    rows = [['keep', 'x'], [], ['keep'], ['hidden', 'z']]
    s = session(rows, header=False)
    s.editor.set_filter(0, query='keep')
    header, _, indices, token = s.view()
    assert header is None and indices == [0, 2]
    s.edit(1, 1, '00123', token)
    assert s.editor.rows == [rows[0], [], ['keep', '00123'], rows[3]]
    assert s.editor.undo() and s.editor.rows == rows


def test_save_filtered_all_rows_and_savepoint_preserves_filter_and_history(tmp_path):
    s = session()
    s.editor.set_filter(2, query='東京')
    token = s.view()[3]
    s.edit(1, 0, '00099', token)
    target = tmp_path/'all.csv'
    s.save_full_copy(target)
    restored = load_csv(target, encoding='utf-8')
    assert restored.rows == s.editor.rows and restored.format == s.format
    assert len(restored.rows) == 6 and restored.rows[2] == ROWS[2]
    assert not s.editor.dirty and s.editor.condition == (2, 'contains', '東京')
    assert s.editor.undo() and s.editor.dirty
    assert s.editor.redo() and not s.editor.dirty


def test_save_zero_matches_keeps_every_row(tmp_path):
    s = session()
    s.editor.set_filter(0, query='none')
    assert s.view()[1] == []
    s.save_full_copy(tmp_path/'all.csv')
    assert load_csv(tmp_path/'all.csv', encoding='utf-8').rows == ROWS


def test_save_failure_preserves_unsaved_data_and_history(tmp_path):
    s = session()
    token = s.view()[3]
    s.edit(1, 0, 'changed', token)
    target = tmp_path/'existing.csv'
    target.write_bytes(b'unchanged')
    with pytest.raises(FileExistsError):
        s.save_full_copy(target)
    assert s.editor.dirty and target.read_bytes() == b'unchanged'
    assert s.editor.undo() and s.editor.rows == ROWS


def test_sort_after_clear_uses_all_rows_and_preserves_heights():
    s = session()
    s.editor.set_row_heights([20, 30, 40, 50, 60, 70])
    s.editor.set_filter(2, query='東京')
    s.view()
    with pytest.raises(ValueError):
        s.editor.sort_rows(1, mode='number')
    s.editor.clear_filter()
    s.editor.sort_rows(1, mode='number')
    _, body, indices, _ = s.view()
    assert indices == [1, 2, 3, 4, 5]
    assert [r[0] for r in body] == ['00005', '00002', '00003', '00123', '00004']
    assert s.editor.row_heights == [20, 70, 40, 50, 30, 60]
    assert s.editor.undo() and s.editor.rows == ROWS


@pytest.mark.parametrize('keysym,keycode,state,system,expected', [
    ('z', 52, 4, 'x11', False), ('Z', 52, 5, 'x11', True),
    ('Z', 90, 4, 'win32', False), ('z', 90, 5, 'win32', True),
    ('??', 90, 4, 'win32', False), ('??', 90, 5, 'win32', True),
    ('??', 90, 4, 'x11', None), ('z', 90, 0, 'win32', None),
    ('Z', 90, 1, 'win32', None), ('a', 65, 4, 'win32', None),
    ('v', 86, 4, 'win32', None),
])
def test_physical_history_key_and_windows_ime_fallback(keysym, keycode, state, system, expected):
    event = SimpleNamespace(keysym=keysym, keycode=keycode, state=state)
    assert history_shortcut(event, system) is expected


class Notice:
    def __init__(self):
        self.text = ''

    def set(self, text):
        self.text = text

    def get(self):
        return self.text


@pytest.mark.parametrize('widget_class', ['Entry', 'TEntry', 'Text', 'TCombobox', 'Spinbox', 'TSpinbox'])
def test_history_route_keeps_input_history_separate(widget_class):
    calls = []
    gui = SimpleNamespace(root=SimpleNamespace(tk=SimpleNamespace(call=lambda *a: 'win32')),
                          history=calls.append, key_notice=Notice())
    event = SimpleNamespace(widget=SimpleNamespace(winfo_class=lambda: widget_class), keysym='z', keycode=90, state=4)
    assert ProjectionGui.route_history_key(gui, event) is None
    assert not calls and '入力欄' in gui.key_notice.get()


@pytest.mark.parametrize('widget_class', ['Canvas', 'TButton', 'Tk', 'Frame'])
@pytest.mark.parametrize('state,redo', [(4, False), (5, True)])
def test_history_route_calls_model_once_and_stops_event(widget_class, state, redo):
    calls = []
    gui = SimpleNamespace(root=SimpleNamespace(tk=SimpleNamespace(call=lambda *a: 'win32')),
                          history=calls.append, key_notice=Notice())
    event = SimpleNamespace(widget=SimpleNamespace(winfo_class=lambda: widget_class), keysym='??', keycode=90, state=state)
    assert ProjectionGui.route_history_key(gui, event) == 'break'
    assert calls == [redo] and '履歴キー受信' in gui.key_notice.get()


def test_history_counts_expose_filter_only_and_paste_undo_redo():
    s = session()
    assert s.history_counts == (0, 0)
    s.editor.set_filter(2, query='東京')
    token = s.view()[3]
    assert s.history_counts == (0, 0)
    assert s.paste(0, 1, [['101', '大阪'], ['202', '京都']], token)
    assert s.history_counts == (1, 0)
    assert s.history() and s.history_counts == (0, 1) and s.editor.rows == ROWS
    assert s.history(True) and s.history_counts == (1, 0)
    assert not s.history(True) and s.history_counts == (1, 0)
