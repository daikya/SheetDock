"""保存前の外部変更検出。更新日時・サイズ・内容ハッシュを比較する。"""

import hashlib
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from .csv_io import load_csv, save_csv
from .encoding import resolve_encoding


class ExternalChangeError(RuntimeError):
    pass


@dataclass(frozen=True)
class Fingerprint:
    mtime_ns: int
    size: int
    digest: str


def fingerprint(path):
    try:
        with Path(path).open('rb') as stream:
            import os
            before = os.fstat(stream.fileno())
            digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            after = os.fstat(stream.fileno())
            if (before.st_mtime_ns, before.st_size) != (after.st_mtime_ns, after.st_size):
                raise ExternalChangeError('確認中にファイルが変更されました。もう一度実行してください。')
            return Fingerprint(after.st_mtime_ns, after.st_size, digest)
    except FileNotFoundError:
        return None


def load_checked(path, *, encoding, delimiter=',', choose_encoding=None):
    before = fingerprint(path)
    if encoding == 'auto':
        encoding = resolve_encoding(Path(path).read_bytes(), choose_encoding)
    document = load_csv(path, encoding=encoding, delimiter=delimiter)
    after = fingerprint(path)
    if before != after:
        raise ExternalChangeError('読込中にファイルが変更されました。もう一度開いてください。')
    return document, after


@contextmanager
def save_lock(path):
    target = Path(path).resolve()
    lock = target.with_name('.' + target.name + '.sheetdock-save.lock')
    try:
        stream = lock.open('x')
    except FileExistsError as exc:
        raise RuntimeError('別のSheetDockが保存中、または前回の保存ロックが残っています。保存を中止しました。') from exc
    try:
        stream.close()
        yield
    finally:
        lock.unlink(missing_ok=True)


def save_checked(document, path, expected, *, backup=True):
    with save_lock(path):
        return _save_locked(document, path, expected, backup=backup)


def _save_locked(document, path, expected, *, backup=True):
    def verify():
        if fingerprint(path) != expected:
            raise ExternalChangeError('ファイルが外部で変更または削除されました。保存を中止しました。')
    verify()
    save_csv(document, path, backup=backup, before_replace=verify)
    # 書込後に外部変更があれば、保存済みにはせず編集内容を保持する。
    restored, stamp = load_checked(path, encoding=document.format.encoding, delimiter=document.format.delimiter)
    if restored != document:
        raise ExternalChangeError('保存直後に内容が変更されました。編集内容は保持しています。')
    return stamp
