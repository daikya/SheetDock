"""Independent CSV row sorting and read-only filter projection; not used by the app."""
from __future__ import annotations

import argparse
import hashlib
import os
from pathlib import Path
import tempfile

from sheetdock.csv_io import CsvDocument, CsvFormat, load_csv, save_csv


from sheetdock.csv_tools import SortFilterEditor


def save_copy(editor, fmt, destination):
    if editor.condition is not None:
        raise ValueError('この検証ではフィルター中の保存は行いません。解除してください')
    target = Path(destination)
    if target.exists():
        raise FileExistsError('出力先が既に存在します。別の名前を指定してください')
    # Validate and publish without replacing an existing file, including creation races.
    with tempfile.TemporaryDirectory(dir=target.parent) as folder:
        candidate = Path(folder) / 'candidate.csv'
        save_csv(CsvDocument(editor.rows, fmt), candidate, backup=False)
        os.link(candidate, target)


def make_samples(folder):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    paths = [folder / name for name in ('sort_basic.csv', 'sort_invalid.csv')]
    if any(path.exists() for path in paths):
        raise FileExistsError('同名サンプルが存在します。別フォルダーを指定してください')
    rows = [['コード', '数量', '区分', 'メモ'],
            ['00123', '10', '東京', '改行\nあり'],
            ['00002', '2', '大阪', '引用符"と,カンマ'],
            ['00003', '2', '東京', '同値の順序を保持'],
            ['00004', '', '京都', '空セル'],
            ['00005', '-1', '東京', '負数']]
    fmt = CsvFormat('utf-8', True, '\r\n', ',')
    for path, data in zip(paths, (rows, rows + [['00006', 'abc', '東京', '数値順を拒否']])):
        save_copy(SortFilterEditor(data), fmt, path)
    return paths


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('path', nargs='?')
    parser.add_argument('--make-samples')
    parser.add_argument('--encoding', default='utf-8', choices=('utf-8', 'cp932'))
    parser.add_argument('--delimiter', default=',', choices=(',', ';', 'tab'))
    parser.add_argument('--no-header', action='store_true')
    parser.add_argument('--sort-column', type=int, help='1始まりの列番号')
    parser.add_argument('--mode', default='text', choices=('text', 'number'))
    parser.add_argument('--descending', action='store_true')
    parser.add_argument('--filter-column', type=int, help='1始まりの列番号')
    parser.add_argument('--operator', default='contains', choices=('contains', 'equals', 'empty', 'nonempty'))
    parser.add_argument('--query', default='')
    parser.add_argument('--save-as')
    args = parser.parse_args(argv)
    try:
        if args.make_samples:
            for path in make_samples(args.make_samples):
                print(path)
            return 0
        if not args.path:
            parser.error('CSVパスまたは--make-samplesが必要です')
        source = Path(args.path)
        original_hash = hashlib.sha256(source.read_bytes()).digest()
        doc = load_csv(source, encoding=args.encoding,
                       delimiter='\t' if args.delimiter == 'tab' else args.delimiter)
        editor = SortFilterEditor(doc.rows, header=not args.no_header)
        before = (editor.rows, editor.row_heights[:])
        if args.sort_column is not None:
            changed = editor.sort_rows(args.sort_column - 1, mode=args.mode, descending=args.descending)
            after = (editor.rows, editor.row_heights[:])
            if changed:
                assert editor.undo() and (editor.rows, editor.row_heights) == before
                assert editor.redo() and (editor.rows, editor.row_heights) == after
            print('並べ替え・Undo/Redoの検証完了（変化なしの場合は履歴追加なし）')
        if args.filter_column is not None:
            editor.set_filter(args.filter_column - 1, operator=args.operator, query=args.query)
        print(f'未保存: {"あり" if editor.dirty else "なし"}')
        rows = editor.rows
        for index in editor.visible_indices():
            print(f'データ行{index + 1}: {rows[index]!r}')
        if args.save_as:
            save_copy(editor, doc.format, args.save_as)
            restored = load_csv(args.save_as, encoding=args.encoding, delimiter=doc.format.delimiter)
            assert restored.rows == editor.rows and restored.format == doc.format
            print(f'全行の別名保存・再読込検証完了: {args.save_as}')
        assert hashlib.sha256(source.read_bytes()).digest() == original_hash
        print('元ファイルは変更していません。本体GUIにはまだ接続していません。')
        return 0
    except (ValueError, TypeError, IndexError, OSError) as error:
        print(f'操作中止：{error}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
