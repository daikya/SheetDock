from types import SimpleNamespace

import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat, load_csv, save_csv
from sheetdock.gui import Session, MultiFileApp


def session(rows, use_header=True):
    return Session(CsvDocument(rows, CsvFormat('utf-8', True, '\r\n', ',')), use_header)


@pytest.mark.parametrize('rows', [[], [[]], [['見出し', 'ID']], [['見出し', 'ID'], ['鈴木', '00123'], [], ['短い行']]])
@pytest.mark.parametrize('initial', [True, False])
def test_toggle_preserves_all_rows_format_and_clean_state(rows, initial):
    s = session(rows, initial)
    original = s.snapshot()
    s.set_header_view(not initial)
    assert s.snapshot() == original
    assert not s.editor.dirty
    assert not s.editor.undo()
    s.set_header_view(initial)
    assert s.snapshot() == original


def test_toggle_preserves_edit_history_redo_and_dirty():
    s = session([['名前', 'ID'], ['鈴木', '00123']])
    s.editor.set_cell(1, 0, '編集')
    s.editor.undo()
    s.set_header_view(False)
    assert s.editor.redo()
    assert s.editor.rows[1][0] == '編集'
    assert s.editor.dirty
    s.set_header_view(True)
    assert s.editor.undo()
    assert s.editor.rows[1][0] == '鈴木'
    assert not s.editor.dirty


@pytest.mark.parametrize('header', [True, False])
def test_edit_first_visible_row_after_toggle_has_correct_mapping(header):
    s = session([['名前', 'ID'], ['鈴木', '00123']], not header)
    s.set_header_view(header)
    _, body = s.view()
    body[0][0] = '変更'
    s.accept_body(body)
    assert s.editor.rows[int(header)][0] == '変更'
    assert s.editor.rows[1][1] == '00123'
    assert s.editor.undo()
    assert s.editor.rows == [['名前', 'ID'], ['鈴木', '00123']]


def test_toggle_is_independent_between_tabs():
    a = session([['見出し'], ['A']])
    b = session([['見出し'], ['B']])
    a.set_header_view(False)
    assert a.view()[0] is None
    assert b.view()[0] == ['見出し']


@pytest.mark.parametrize('header', [True, False])
def test_toggle_and_save_preserve_first_row(tmp_path, header):
    s = session([['名前', 'ID'], ['鈴木', '00123']], not header)
    s.set_header_view(header)
    path = tmp_path / 'saved.csv'
    save_csv(s.snapshot(), path)
    assert load_csv(path, encoding='utf-8').rows == [['名前', 'ID'], ['鈴木', '00123']]


def test_gui_commits_pending_cell_before_header_switch():
    app = MultiFileApp.__new__(MultiFileApp)
    app.active = object()
    app.session = session([['名前'], ['鈴木']])
    selected = [False]
    app.header_display = SimpleNamespace(get=lambda: selected[0], set=lambda value: selected.__setitem__(0, value))
    events = []
    def snapshot():
        events.append(('snapshot', app.session.has_header))
        app.session.accept_body([['編集途中']])
        # 編集通知でチェック状態が旧モードに更新されても、要求は先に取得する。
        selected[0] = True
    app.snapshot = snapshot
    app.render = lambda: events.append(('render', app.session.has_header))
    app.toggle_header()
    assert events == [('snapshot', True), ('render', False)]
    assert app.session.editor.rows == [['名前'], ['編集途中']]
    assert app.session.editor.undo()
    assert app.session.editor.rows == [['名前'], ['鈴木']]
