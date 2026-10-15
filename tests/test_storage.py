import os

import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat, save_csv
from sheetdock.storage import ExternalChangeError, fingerprint, load_checked, save_checked, save_lock


def fixture(tmp_path):
    path = tmp_path / 'data.csv'
    doc = CsvDocument([['ID'], ['00123']], CsvFormat('utf-8', True, '\r\n', ','))
    save_csv(doc, path)
    loaded, stamp = load_checked(path, encoding='utf-8')
    return path, loaded, stamp


def test_normal_save_and_backup(tmp_path):
    path, doc, stamp = fixture(tmp_path)
    original = path.read_bytes()
    doc.rows[1][0] = '00999'
    saved = save_checked(doc, path, stamp)
    assert saved == fingerprint(path)
    assert saved != stamp
    assert path.with_name('data.csv.bak').read_bytes() == original
    assert load_checked(path, encoding='utf-8')[0] == doc


@pytest.mark.parametrize('mode', ['modified', 'deleted', 'same_size_and_time'])
def test_external_change_is_detected_without_overwrite(tmp_path, mode):
    path, doc, stamp = fixture(tmp_path)
    previous_stat = path.stat()
    if mode == 'deleted':
        path.unlink()
        external = None
    else:
        path.write_bytes(path.read_bytes().replace(b'00123', b'00555'))
        if mode == 'same_size_and_time':
            os.utime(path, ns=(previous_stat.st_atime_ns, previous_stat.st_mtime_ns))
        external = path.read_bytes()
    doc.rows[1][0] = '00999'
    with pytest.raises(ExternalChangeError):
        save_checked(doc, path, stamp)
    assert (path.read_bytes() if path.exists() else None) == external
    assert doc.rows[1][0] == '00999'
    assert not path.with_name('data.csv.bak').exists()


def test_explicit_overwrite_backs_up_latest_external_content(tmp_path):
    path, doc, stamp = fixture(tmp_path)
    path.write_bytes(path.read_bytes().replace(b'00123', b'00555'))
    external = path.read_bytes()
    current = fingerprint(path)
    doc.rows[1][0] = '00999'
    save_checked(doc, path, current)
    assert path.with_name('data.csv.bak').read_bytes() == external
    assert load_checked(path, encoding='utf-8')[0].rows[1][0] == '00999'


def test_explicit_recreate_deleted_file(tmp_path):
    path, doc, stamp = fixture(tmp_path)
    path.unlink()
    save_checked(doc, path, None)
    assert path.exists()


def test_change_during_backup_is_detected_before_replace(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_io
    path, doc, stamp = fixture(tmp_path)
    copy2 = csv_io.shutil.copy2
    def alter(source, target):
        copy2(source, target)
        path.write_bytes(path.read_bytes().replace(b'00123', b'00555'))
    monkeypatch.setattr(csv_io.shutil, 'copy2', alter)
    doc.rows[1][0] = '00999'
    with pytest.raises(ExternalChangeError):
        save_checked(doc, path, stamp)
    assert b'00555' in path.read_bytes()
    assert not list(tmp_path.glob('*.tmp'))


def test_read_race_rejected(tmp_path, monkeypatch):
    import sheetdock.storage as storage
    path, doc, stamp = fixture(tmp_path)
    load = storage.load_csv
    def alter(*args, **kwargs):
        result = load(*args, **kwargs)
        path.write_bytes(path.read_bytes().replace(b'00123', b'00555'))
        return result
    monkeypatch.setattr(storage, 'load_csv', alter)
    with pytest.raises(ExternalChangeError):
        storage.load_checked(path, encoding='utf-8')


def test_replace_failure_keeps_original_and_cleans_temp(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_io
    path, doc, stamp = fixture(tmp_path)
    original = path.read_bytes()
    def fail(*args):
        raise PermissionError('Excelなどで使用中')
    monkeypatch.setattr(csv_io.os, 'replace', fail)
    doc.rows[1][0] = '00999'
    with pytest.raises(PermissionError):
        save_checked(doc, path, stamp)
    assert path.read_bytes() == original
    assert doc.rows[1][0] == '00999'
    assert not list(tmp_path.glob('*.tmp'))


def test_save_as_new_path_detects_newly_created_target(tmp_path):
    path, doc, stamp = fixture(tmp_path)
    target = tmp_path / 'new.csv'
    expected = fingerprint(target)
    target.write_bytes(b'external')
    with pytest.raises(ExternalChangeError):
        save_checked(doc, target, expected)
    assert target.read_bytes() == b'external'


def test_concurrent_sheetdock_save_is_rejected_and_lock_released(tmp_path):
    path, doc, stamp = fixture(tmp_path)
    original = path.read_bytes()
    with save_lock(path):
        with pytest.raises(RuntimeError, match='保存中'):
            save_checked(doc, path, stamp)
        assert path.read_bytes() == original
    assert not list(tmp_path.glob('*.sheetdock-save.lock'))
    assert save_checked(doc, path, stamp) == fingerprint(path)


def test_failed_save_releases_lock(tmp_path):
    path, doc, stamp = fixture(tmp_path)
    path.write_bytes(path.read_bytes().replace(b'00123', b'00555'))
    with pytest.raises(ExternalChangeError):
        save_checked(doc, path, stamp)
    assert not list(tmp_path.glob('*.sheetdock-save.lock'))
