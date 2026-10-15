import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat, save_csv, load_csv
from sheetdock.gui import Session


def session(header=True):
    return Session(CsvDocument([['名前', 'ID'], ['鈴木', '00123'], ['佐藤', '00000']], CsvFormat('utf-8', True, '\r\n', ',')), header)


@pytest.mark.parametrize('header', [True, False])
def test_widget_edit_and_duplicate_notification_record_one_action(header):
    s = session(header)
    _, body = s.view()
    body[-1][0] = '変更'
    assert s.accept_body(body)
    assert not s.accept_body(body)
    assert s.editor.undo()
    assert s.editor.rows[-1][0] == '佐藤'
    assert not s.editor.undo()


def test_header_column_insert_and_body_edit_share_history():
    s = session()
    s.editor.insert_columns(1, 2)
    header, body = s.view()
    assert header == ['名前', '', '', 'ID']
    body[0][1] = '追加'
    s.accept_body(body)
    s.editor.undo()
    assert s.view()[1][0] == ['鈴木', '', '', '00123']
    s.editor.undo()
    assert s.editor.rows[0] == ['名前', 'ID']


def test_body_row_selection_does_not_delete_header():
    s = session()
    selected = [0, 1]
    s.editor.delete_rows([r + s.offset for r in selected])
    assert s.editor.rows == [['名前', 'ID']]
    s.editor.undo()
    assert len(s.view()[1]) == 2


def test_invalid_widget_value_is_atomic():
    s = session()
    original = s.editor.rows
    with pytest.raises(TypeError):
        s.accept_body([[123]])
    assert s.editor.rows == original
    assert not s.editor.undo()


def test_combined_history_saves_header_and_body(tmp_path):
    s = session()
    s.editor.set_cell(0, 0, '氏名')
    _, body = s.view()
    body[0][0] = '編集'
    s.accept_body(body)
    target = tmp_path / 'saved.csv'
    save_csv(s.snapshot(), target)
    s.editor.mark_saved()
    assert not s.editor.dirty
    assert load_csv(target, encoding='utf-8').rows == [['氏名', 'ID'], ['編集', '00123'], ['佐藤', '00000']]
    s.editor.undo()
    assert s.editor.dirty
    s.editor.undo()
    assert s.editor.rows[0][0] == '名前'
