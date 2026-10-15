"""CSVの読込・検証・安全な書込。文字コードは明示指定。"""

import codecs
import csv
import io
import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CsvFormat:
    encoding: str
    bom: bool
    newline: str
    delimiter: str


@dataclass
class CsvDocument:
    rows: list[list[str]]
    format: CsvFormat


def _record_newline(text: str, delimiter: str) -> str:
    """引用符内のセル内改行を除き、レコード区切りを調べる。"""
    found = set()
    quoted = False
    field_start = True
    i = 0
    while i < len(text):
        char = text[i]
        if quoted:
            if char == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    i += 2
                    continue
                quoted = False
        elif char == '"' and field_start:
            quoted = True
            field_start = False
        elif char in '\r\n':
            ending = '\r\n' if text[i:i + 2] == '\r\n' else char
            found.add(ending)
            field_start = True
            i += len(ending)
            continue
        elif char == delimiter:
            field_start = True
        else:
            field_start = False
        i += 1
    if len(found) > 1:
        raise ValueError('レコードの改行コードが混在しています。正規化方針が未確定のため保存検証対象外です。')
    return next(iter(found), '\r\n')


def load_csv(path: str | Path, *, encoding: str, delimiter: str = ',') -> CsvDocument:
    """文字コードと区切り文字は明示する。推測による誤読を避ける。"""
    if delimiter not in (',', '\t', ';'):
        raise ValueError('検証対象の区切り文字はカンマ・タブ・セミコロンです。')
    normalized = codecs.lookup(encoding).name
    if normalized == 'utf-8-sig':
        normalized = 'utf-8'
    if normalized not in ('utf-8', 'cp932'):
        raise ValueError('検証対象の文字コードはUTF-8とCP932です。')
    raw = Path(path).read_bytes()
    bom = raw.startswith(codecs.BOM_UTF8)
    if bom and normalized != 'utf-8':
        raise ValueError('UTF-8 BOMと指定文字コードが一致しません。')
    text = (raw[len(codecs.BOM_UTF8):] if bom else raw).decode(normalized, errors='strict')
    rows = list(csv.reader(io.StringIO(text, newline=''), delimiter=delimiter, strict=True))
    return CsvDocument(rows, CsvFormat(normalized, bom, _record_newline(text, delimiter), delimiter))


def save_csv(document: CsvDocument, path: str | Path, *, backup: bool = True, before_replace=None) -> None:
    """同じフォルダーに一時保存し、再読込一致を確認してから置換する。"""
    target = Path(path)
    fmt = document.format
    if any(not isinstance(cell, str) for row in document.rows for cell in row):
        raise TypeError('CSVのセルはすべて文字列でなければなりません。')
    buffer = io.StringIO(newline='')
    # LF形式でもセル内の単独CRを確実に引用する。元の引用スタイルは保持対象外。
    csv.writer(buffer, delimiter=fmt.delimiter, lineterminator=fmt.newline, quoting=csv.QUOTE_ALL).writerows(document.rows)
    payload = buffer.getvalue().encode(fmt.encoding, errors='strict')
    if fmt.bom:
        if codecs.lookup(fmt.encoding).name != 'utf-8':
            raise ValueError('BOMはUTF-8のみ検証対象です。')
        payload = codecs.BOM_UTF8 + payload
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=target.parent, prefix=f'.{target.name}.', suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        restored = load_csv(temporary, encoding=fmt.encoding, delimiter=fmt.delimiter)
        if restored.rows != document.rows or restored.format != fmt:
            raise ValueError('一時ファイルの再読込検証に失敗しました。')
        if before_replace is not None:
            before_replace()
        if backup and target.exists():
            # 技術検証では1世代のみ。本体の命名・世代管理は別途決める。
            shutil.copy2(target, target.with_name(target.name + '.bak'))
        if before_replace is not None:
            before_replace()
        os.replace(temporary, target)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
