"""Experimental XLSX safe save, reusing the CSV application's fingerprints/lock."""
from pathlib import Path
import os
import shutil
import tempfile

from openpyxl import load_workbook
from probes.xlsx_probe import XlsxProbe, snapshot
from sheetdock.storage import ExternalChangeError, fingerprint, save_lock


def load_checked(path):
    before = fingerprint(path)
    probe = XlsxProbe(path)
    after = fingerprint(path)
    if before is None or before != after:
        probe.book.close()
        raise ExternalChangeError('読込中にファイルが変更されました。もう一度開いてください。')
    return probe, after


def save_checked(probe, path, expected, *, backup=True):
    if probe.read_only:
        raise PermissionError('読み取り専用のブックは保存できません')
    target = Path(path).resolve()
    if target.suffix.lower() != '.xlsx':
        raise ValueError('保存先は.xlsxを指定してください')
    temporary, backup_temp = None, None

    def verify():
        if fingerprint(target) != expected:
            raise ExternalChangeError('ファイルが外部で変更または削除されました。保存を中止しました。')

    with save_lock(target):
        try:
            verify()
            with tempfile.NamedTemporaryFile(dir=target.parent, prefix=f'.{target.name}.',
                                             suffix='.xlsx', delete=False) as stream:
                temporary = Path(stream.name)
                probe.book.save(stream)
                stream.flush()
                os.fsync(stream.fileno())
            checked = load_workbook(temporary, data_only=False)
            try:
                if snapshot(checked) != snapshot(probe.book):
                    raise ValueError('一時ファイルの再読込検証に失敗しました')
            finally:
                checked.close()
            written = fingerprint(temporary)
            verify()
            if backup and expected is not None:
                with tempfile.NamedTemporaryFile(dir=target.parent, prefix=f'.{target.name}.backup.',
                                                 suffix='.tmp', delete=False) as stream:
                    backup_temp = Path(stream.name)
                shutil.copy2(target, backup_temp)
                old = fingerprint(backup_temp)
                if (old.size, old.digest) != (expected.size, expected.digest):
                    raise ExternalChangeError('バックアップ中にファイルが変更されました')
                verify()
                os.replace(backup_temp, target.with_name(target.name + '.bak'))
            verify()
            if expected is None:
                # Atomic publication without overwriting a concurrently created file.
                # If the filesystem cannot hard-link, fail safely and keep edits.
                os.link(temporary, target)
            else:
                os.replace(temporary, target)
            after = fingerprint(target)
            if after is None or (after.size, after.digest) != (written.size, written.digest):
                raise ExternalChangeError('保存直後にファイルが変更されました。編集内容は保持しています。')
            return after
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
            if backup_temp is not None:
                backup_temp.unlink(missing_ok=True)
