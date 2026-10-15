"""実行: python -m pytest tests -q"""

import codecs
import copy
import csv
import io

import pytest

from sheetdock.csv_io import load_csv, save_csv


ROWS = [
    ['ID', '日本語', 'メモ', '空欄'],
    ['00123', '鈴木', 'カンマ,と"引用符"', ''],
    ['00000', '改行', '一行目\n二行目', ''],
    ['1E10', '2026-10-13', 'タブ\tとCRLF\r\nあり', ''],
    ['短い行'],
    [],
]


def write_fixture(path, *, encoding='utf-8', bom=False, newline='\r\n', delimiter=',', rows=None):
    stream = io.StringIO(newline='')
    csv.writer(stream, delimiter=delimiter, lineterminator=newline).writerows(ROWS if rows is None else rows)
    data = stream.getvalue().encode(encoding)
    path.write_bytes((codecs.BOM_UTF8 if bom else b'') + data)


@pytest.mark.parametrize('encoding,bom', [('utf-8', False), ('utf-8', True), ('cp932', False)])
@pytest.mark.parametrize('newline', ['\r\n', '\n'])
@pytest.mark.parametrize('delimiter', [',', '\t', ';'])
def test_edit_and_roundtrip_preserves_values_and_format(tmp_path, encoding, bom, newline, delimiter):
    source = tmp_path / 'data.csv'
    write_fixture(source, encoding=encoding, bom=bom, newline=newline, delimiter=delimiter)
    original = source.read_bytes()
    doc = load_csv(source, encoding=encoding, delimiter=delimiter)
    assert doc.rows == ROWS
    assert doc.format.newline == newline
    assert doc.format.bom == bom
    doc.rows[1][1] = '編集済み'
    expected = copy.deepcopy(doc)
    save_csv(doc, source)
    assert load_csv(source, encoding=encoding, delimiter=delimiter) == expected
    assert source.with_name('data.csv.bak').read_bytes() == original
    assert source.read_bytes().startswith(codecs.BOM_UTF8) == bom
    assert not list(tmp_path.glob('*.tmp'))


def test_no_backup_when_disabled(tmp_path):
    path = tmp_path / 'data.csv'
    write_fixture(path)
    save_csv(load_csv(path, encoding='utf-8'), path, backup=False)
    assert not path.with_name('data.csv.bak').exists()


def test_replace_failure_protects_original_and_edit(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_probe
    path = tmp_path / 'data.csv'
    write_fixture(path)
    original = path.read_bytes()
    doc = load_csv(path, encoding='utf-8')
    doc.rows[1][0] = '00999'
    expected = copy.deepcopy(doc)
    def fail(*args):
        raise PermissionError('ファイルが使用中')
    monkeypatch.setattr(csv_probe.os, 'replace', fail)
    with pytest.raises(PermissionError):
        save_csv(doc, path)
    assert path.read_bytes() == original
    assert doc == expected
    assert not list(tmp_path.glob('*.tmp'))


def test_encoding_failure_does_not_touch_source(tmp_path):
    path = tmp_path / 'data.csv'
    write_fixture(path, encoding='cp932')
    original = path.read_bytes()
    doc = load_csv(path, encoding='cp932')
    doc.rows[1][1] = '😀'
    with pytest.raises(UnicodeEncodeError):
        save_csv(doc, path)
    assert path.read_bytes() == original
    assert doc.rows[1][1] == '😀'
    assert not list(tmp_path.glob('*.tmp'))


def test_backup_failure_does_not_replace_source(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_probe
    path = tmp_path / 'data.csv'
    write_fixture(path)
    original = path.read_bytes()
    doc = load_csv(path, encoding='utf-8')
    doc.rows[1][1] = '編集'
    def fail(*args):
        raise OSError('バックアップ失敗')
    monkeypatch.setattr(csv_probe.shutil, 'copy2', fail)
    with pytest.raises(OSError):
        save_csv(doc, path)
    assert path.read_bytes() == original
    assert doc.rows[1][1] == '編集'
    assert not list(tmp_path.glob('*.tmp'))


def test_verification_failure_does_not_replace_source(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_probe
    path = tmp_path / 'data.csv'
    write_fixture(path)
    original = path.read_bytes()
    doc = load_csv(path, encoding='utf-8')
    def wrong_read(*args, **kwargs):
        result = copy.deepcopy(doc)
        result.rows[1][0] = '異なる値'
        return result
    monkeypatch.setattr(csv_probe, 'load_csv', wrong_read)
    with pytest.raises(ValueError, match='再読込検証'):
        save_csv(doc, path)
    assert path.read_bytes() == original
    assert not list(tmp_path.glob('*.tmp'))


def test_non_string_value_rejected(tmp_path):
    path = tmp_path / 'data.csv'
    write_fixture(path)
    original = path.read_bytes()
    doc = load_csv(path, encoding='utf-8')
    doc.rows[1][0] = 123
    with pytest.raises(TypeError):
        save_csv(doc, path)
    assert path.read_bytes() == original


def test_mixed_record_newlines_rejected(tmp_path):
    path = tmp_path / 'mixed.csv'
    path.write_bytes(b'a,b\r\nc,d\n')
    with pytest.raises(ValueError, match='混在'):
        load_csv(path, encoding='utf-8')


def test_bom_encoding_mismatch_rejected(tmp_path):
    path = tmp_path / 'bom.csv'
    write_fixture(path, bom=True)
    with pytest.raises(ValueError, match='BOM'):
        load_csv(path, encoding='cp932')


def test_empty_file_roundtrip(tmp_path):
    path = tmp_path / 'empty.csv'
    path.write_bytes(b'')
    doc = load_csv(path, encoding='utf-8')
    save_csv(doc, path)
    assert path.read_bytes() == b''


def test_wrong_encoding_is_not_silently_replaced(tmp_path):
    path = tmp_path / 'data.csv'
    write_fixture(path, encoding='cp932')
    with pytest.raises(UnicodeDecodeError):
        load_csv(path, encoding='utf-8')


def test_lf_file_preserves_new_bare_cr_inside_cell(tmp_path):
    path = tmp_path / 'data.csv'
    write_fixture(path, newline='\n')
    doc = load_csv(path, encoding='utf-8')
    doc.rows[1][2] = '単独\r改行'
    save_csv(doc, path)
    assert load_csv(path, encoding='utf-8') == doc


def test_non_delimiter_before_quote_is_not_a_quoted_field(tmp_path):
    path = tmp_path / 'data.csv'
    path.write_bytes(b'a;"x,b\nc,d\n')
    doc = load_csv(path, encoding='utf-8', delimiter=',')
    assert doc.format.newline == '\n'
    assert doc.rows[0][0] == 'a;"x'
    save_csv(doc, path)
    assert load_csv(path, encoding='utf-8') == doc


def test_flush_failure_keeps_original_and_cleans_temporary(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_probe
    path = tmp_path / 'data.csv'
    write_fixture(path)
    original = path.read_bytes()
    doc = load_csv(path, encoding='utf-8')
    def fail(*args):
        raise OSError('ディスク書込失敗')
    monkeypatch.setattr(csv_probe.os, 'fsync', fail)
    with pytest.raises(OSError):
        save_csv(doc, path)
    assert path.read_bytes() == original
    assert not list(tmp_path.glob('*.tmp'))


def test_save_as_new_file(tmp_path):
    source = tmp_path / 'source.csv'
    target = tmp_path / 'new.csv'
    write_fixture(source)
    original = source.read_bytes()
    doc = load_csv(source, encoding='utf-8')
    save_csv(doc, target)
    assert source.read_bytes() == original
    assert load_csv(target, encoding='utf-8') == doc
    assert not target.with_name('new.csv.bak').exists()
