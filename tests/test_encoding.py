import codecs

import pytest

from sheetdock.encoding import detect_encoding, resolve_encoding, AmbiguousEncodingError, EncodingSelectionCancelled
from sheetdock.storage import load_checked, save_checked, fingerprint, ExternalChangeError


def test_utf8_bom_has_priority():
    result = detect_encoding(codecs.BOM_UTF8 + '名前,ID\r\n鈴木,00123\r\n'.encode('utf-8'))
    assert result.selected == 'utf-8'
    assert result.candidates == ('utf-8',)


def test_cp932_only_is_detected():
    assert detect_encoding('名前,鈴木'.encode('cp932')).selected == 'cp932'


def test_utf8_only_is_detected():
    assert detect_encoding('あ'.encode('utf-8')).selected == 'utf-8'


@pytest.mark.parametrize('raw', [b'', b'ID\r\n00123\r\n'])
def test_identical_ascii_uses_utf8_without_prompt(raw):
    result = detect_encoding(raw)
    assert result.selected == 'utf-8'
    assert result.candidates == ('utf-8', 'cp932')


def test_ambiguous_text_requires_choice():
    raw = '表'.encode('utf-8')
    result = detect_encoding(raw)
    assert result.selected is None
    assert result.previews['utf-8'] != result.previews['cp932']
    with pytest.raises(AmbiguousEncodingError):
        resolve_encoding(raw)
    assert resolve_encoding(raw, lambda decision: 'utf-8') == 'utf-8'
    assert resolve_encoding(raw, lambda decision: 'cp932') == 'cp932'


def test_choice_cancelled():
    with pytest.raises(EncodingSelectionCancelled):
        resolve_encoding('表'.encode('utf-8'), lambda d: None)


def test_invalid_choice_rejected():
    with pytest.raises(ValueError):
        resolve_encoding('表'.encode('utf-8'), lambda d: 'latin1')


@pytest.mark.parametrize('raw', [b'\x81', b'abc\x00def', '名前'.encode('utf-16'), '名前'.encode('utf-32')])
def test_unsupported_or_binary_data_rejected(raw):
    with pytest.raises(ValueError):
        detect_encoding(raw)


def test_invalid_utf8_bom_payload_is_not_guessed_as_cp932():
    with pytest.raises(UnicodeDecodeError):
        detect_encoding(codecs.BOM_UTF8 + b'\xff')


@pytest.mark.parametrize('encoding,bom', [('utf-8', False), ('utf-8', True), ('cp932', False)])
def test_auto_read_edit_save_preserves_encoding_and_bom(tmp_path, encoding, bom):
    path = tmp_path / 'data.csv'
    raw = '名前,ID\r\n鈴木,00123\r\n'.encode(encoding)
    path.write_bytes((codecs.BOM_UTF8 if bom else b'') + raw)
    doc, stamp = load_checked(path, encoding='auto', choose_encoding=lambda d: encoding)
    assert doc.rows[1] == ['鈴木', '00123']
    assert doc.format.encoding == encoding
    assert doc.format.bom == bom
    doc.rows[1][0] = '編集'
    save_checked(doc, path, stamp)
    restored, _ = load_checked(path, encoding=encoding)
    assert restored == doc
    assert path.read_bytes().startswith(codecs.BOM_UTF8) == bom


def test_manual_read_bypasses_choice(tmp_path):
    path = tmp_path / 'data.csv'
    path.write_bytes('表,00123\n'.encode('utf-8'))
    def unexpected(d):
        raise AssertionError('手動指定で選択画面は不要')
    doc, _ = load_checked(path, encoding='utf-8', choose_encoding=unexpected)
    assert doc.rows == [['表', '00123']]


def test_change_while_choosing_is_rejected(tmp_path):
    path = tmp_path / 'data.csv'
    path.write_bytes('表,00123\n'.encode('utf-8'))
    def choose(d):
        path.write_bytes('表,00999\n'.encode('utf-8'))
        return 'utf-8'
    with pytest.raises(ExternalChangeError):
        load_checked(path, encoding='auto', choose_encoding=choose)


def test_cancel_does_not_change_file(tmp_path):
    path = tmp_path / 'data.csv'
    raw = '表,00123\n'.encode('utf-8')
    path.write_bytes(raw)
    with pytest.raises(EncodingSelectionCancelled):
        load_checked(path, encoding='auto', choose_encoding=lambda d: None)
    assert path.read_bytes() == raw
