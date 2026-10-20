from probes.xlsx_inline_probe import main


def test_samples_never_overwrite_existing_files(tmp_path):
    assert main(['--make-samples', str(tmp_path)]) == 0
    before = (tmp_path / 'inline_basic.xlsx').read_bytes()
    assert main(['--make-samples', str(tmp_path)]) == 1
    assert (tmp_path / 'inline_basic.xlsx').read_bytes() == before
