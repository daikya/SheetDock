"""GUI接続部のデータ保護テスト。画面操作はWindows実機で別途確認。"""

import copy

import pytest

from csv_probe import CsvDocument, CsvFormat, load_csv, save_csv
from csv_gui_probe import split_rows, build_document


@pytest.mark.parametrize('rows', [[], [[]], [['名前', 'ID']], [['名前', 'ID'], ['鈴木', '00123'], [], ['末尾']]])
@pytest.mark.parametrize('use_header', [True, False])
def test_header_split_and_restore_does_not_change_data(rows, use_header):
    original = copy.deepcopy(rows)
    doc = CsvDocument(rows, CsvFormat('utf-8', True, '\r\n', ','))
    header, body = split_rows(doc.rows, use_header)
    rebuilt = build_document(doc, header, body)
    assert rebuilt == doc
    assert doc.rows == original
    assert rebuilt.rows is not doc.rows


def test_edit_body_preserves_header_and_original(tmp_path):
    doc = CsvDocument([['名前', 'ID'], ['鈴木', '00123']], CsvFormat('utf-8', True, '\r\n', ','))
    original = copy.deepcopy(doc)
    header, body = split_rows(doc.rows, True)
    body[0][0] = '編集済み'
    edited = build_document(doc, header, body)
    path = tmp_path / 'saved.csv'
    save_csv(edited, path)
    assert load_csv(path, encoding='utf-8') == edited
    assert edited.rows[0] == ['名前', 'ID']
    assert edited.rows[1] == ['編集済み', '00123']
    assert doc == original


def test_snapshot_is_independent_of_widget_data():
    doc = CsvDocument([['見出し'], ['値']], CsvFormat('utf-8', False, '\n', ','))
    header, body = split_rows(doc.rows, True)
    saved = build_document(doc, header, body)
    header[0] = '変更'
    body[0][0] = '変更'
    assert saved.rows == [['見出し'], ['値']]


def test_numeric_cell_is_rejected():
    doc = CsvDocument([], CsvFormat('utf-8', False, '\n', ','))
    with pytest.raises(TypeError):
        build_document(doc, None, [[123]])
