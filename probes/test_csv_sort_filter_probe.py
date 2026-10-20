from copy import deepcopy

import pytest

from probes.csv_sort_filter_probe import SortFilterEditor, main, make_samples, save_copy
from sheetdock.csv_io import CsvFormat, load_csv


ROWS = [['ID', '数量'], ['00123', '10'], ['00002', '2'],
        ['00003', '2'], ['00004', ''], ['00005', '-1']]


def test_active_filter_blocks_sort_and_save_without_changes(tmp_path):
    editor = SortFilterEditor(ROWS)
    editor.set_filter(1, query='2')
    with pytest.raises(ValueError, match='解除'):
        editor.sort_rows(1)
    with pytest.raises(ValueError, match='解除'):
        save_copy(editor, CsvFormat('utf-8', False, '\n', ','), tmp_path/'copy.csv')
    assert editor.rows == ROWS and not editor.dirty
    assert not (tmp_path/'copy.csv').exists()


@pytest.mark.parametrize('fmt', [CsvFormat('utf-8', True, '\r\n', ','),
                                CsvFormat('cp932', False, '\n', '\t')])
def test_safe_new_copy_preserves_format_and_cells(tmp_path, fmt):
    rows = [['名前', '数量'], ['改行\n引用符"と,カンマ', '10'], ['00123', '2']]
    editor = SortFilterEditor(rows)
    editor.sort_rows(1, mode='number')
    path = tmp_path/'copy.csv'
    save_copy(editor, fmt, path)
    restored = load_csv(path, encoding=fmt.encoding, delimiter=fmt.delimiter)
    assert restored.rows == editor.rows and restored.format == fmt
    raw = path.read_bytes()
    with pytest.raises(FileExistsError):
        save_copy(editor, fmt, path)
    assert path.read_bytes() == raw


def test_cli_sort_filter_rejection_source_and_existing_output(tmp_path, capsys):
    assert main(['--make-samples', str(tmp_path)]) == 0
    source = tmp_path/'sort_basic.csv'
    raw = source.read_bytes()
    target = tmp_path/'sorted.csv'
    assert main([str(source), '--sort-column', '2', '--mode', 'number', '--save-as', str(target)]) == 0
    assert load_csv(target, encoding='utf-8').rows[1][0] == '00005'
    assert '再読込検証完了' in capsys.readouterr().out
    assert main([str(source), '--filter-column', '3', '--query', '東京']) == 0
    output = capsys.readouterr().out
    assert '未保存: なし' in output and '大阪' not in output
    assert main([str(source), '--save-as', str(source)]) == 1
    assert main([str(tmp_path/'sort_invalid.csv'), '--sort-column', '2', '--mode', 'number']) == 1
    assert source.read_bytes() == raw
    with pytest.raises(FileExistsError):
        make_samples(tmp_path)
