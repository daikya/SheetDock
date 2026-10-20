import pytest

from probes.xlsx_append_probe import main, make_samples


def test_cli_success_samples_and_existing_output_are_safe(tmp_path, capsys):
    folder = tmp_path / 'samples'
    assert main(['--make-samples', str(folder)]) == 0
    path = folder / 'append_basic.xlsx'
    original = path.read_bytes()
    target = tmp_path / 'copy.xlsx'
    args = [str(path), '--axis', 'rows', '--count', '2', '--save-as', str(target)]
    assert main(args) == 0
    saved = target.read_bytes()
    assert '検証完了' in capsys.readouterr().out
    assert main(args) == 1 and target.read_bytes() == saved and path.read_bytes() == original
    with pytest.raises(FileExistsError):
        make_samples(folder)
    assert path.read_bytes() == original


@pytest.mark.parametrize('kind', ['formula', 'merged', 'readonly'])
def test_cli_restriction_produces_no_output(tmp_path, kind):
    make_samples(tmp_path)
    path = tmp_path / f'append_{kind}.xlsx'
    original = path.read_bytes()
    target = tmp_path / 'rejected.xlsx'
    assert main([str(path), '--axis', 'columns', '--save-as', str(target)]) == 1
    assert not target.exists() and path.read_bytes() == original
