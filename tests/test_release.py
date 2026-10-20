import hashlib
import json
from pathlib import Path
import tomllib
from zipfile import ZipFile

import pytest

import build_release as release
from sheetdock import __version__


def bundle_fixture(tmp_path):
    bundle = tmp_path / 'bundle'
    bundle.mkdir()
    (bundle / 'SheetDock.exe').write_bytes(b'MZ-test-fixture')
    (bundle / '_internal').mkdir()
    (bundle / '_internal' / 'python312.dll').write_bytes(b'test-fixture')
    for name in release.REQUIRED_DOCS:
        (bundle / name).write_text('test document')
    (bundle / 'BUILD_INFO.json').write_text(json.dumps({'sheetdock': '1.0.0'}))
    for group in ('Python', 'Tcl', 'Tk', 'tksheet', 'pyinstaller', 'openpyxl', 'et-xmlfile'):
        folder = bundle / 'licenses' / group
        folder.mkdir(parents=True)
        (folder / 'LICENSE.txt').write_text('test license fixture')
    (bundle / 'licenses' / 'INDEX.txt').write_text('test index')
    return bundle


def test_version_and_mit_license_consistency():
    root = Path(__file__).resolve().parents[1]
    metadata = tomllib.loads((root / 'pyproject.toml').read_text(encoding='utf-8'))
    assert metadata['project']['version'] == __version__
    assert 'MIT License' in (root / 'LICENSE').read_text()


def test_zip_contains_complete_folder_and_checksum(tmp_path):
    bundle = bundle_fixture(tmp_path)
    output = tmp_path / 'release.zip'
    release.create_zip(bundle, output, '1.0.0')
    with ZipFile(output) as archive:
        names = archive.namelist()
        assert 'SheetDock-1.0.0-windows-x64/SheetDock.exe' in names
        assert 'SheetDock-1.0.0-windows-x64/_internal/python312.dll' in names
        assert 'SheetDock-1.0.0-windows-x64/LICENSE' in names
        assert archive.testzip() is None
    assert output.with_suffix('.zip.sha256').read_text().startswith(hashlib.sha256(output.read_bytes()).hexdigest())


@pytest.mark.parametrize('relative', ['settings.json', 'private.csv', 'private.tsv', 'private.xlsx', 'logs/a.log', 'tests/test_private.py', '.venv312/pyvenv.cfg'])
def test_user_or_test_data_blocks_zip_and_preserves_previous_output(tmp_path, relative):
    bundle = bundle_fixture(tmp_path)
    private = bundle / relative
    private.parent.mkdir(parents=True, exist_ok=True)
    private.write_text('private data')
    output = tmp_path / 'release.zip'
    output.write_bytes(b'previous release')
    with pytest.raises(ValueError):
        release.create_zip(bundle, output, '1.0.0')
    assert output.read_bytes() == b'previous release'


def test_missing_license_blocks_zip(tmp_path):
    bundle = bundle_fixture(tmp_path)
    (bundle / 'licenses' / 'Tk' / 'LICENSE.txt').unlink()
    with pytest.raises(ValueError, match='Tk'):
        release.create_zip(bundle, tmp_path / 'release.zip', '1.0.0')


def test_collects_actual_provided_license_bytes(tmp_path):
    home = tmp_path / 'python'
    home.mkdir()
    (home / 'LICENSE.txt').write_text('Python source license')
    for folder in ('tcl8.6', 'tk8.6'):
        path = home / 'tcl' / folder
        path.mkdir(parents=True)
        (path / 'license.terms').write_text(folder + ' original license')
    packages = {}
    for name in ('tksheet', 'pyinstaller', 'openpyxl', 'et-xmlfile'):
        path = tmp_path / (name + '.txt')
        path.write_text(name + ' original license')
        packages[name] = [path]
    bundle = tmp_path / 'bundle'
    release.collect_licenses(bundle, python_home=home, package_paths=packages)
    assert (bundle / 'licenses' / 'Tk' / '01-license.terms').read_text() == 'tk8.6 original license'
    assert (bundle / 'licenses' / 'tksheet' / '01-tksheet.txt').read_text() == 'tksheet original license'
    assert str(tmp_path) not in (bundle / 'licenses' / 'INDEX.txt').read_text()


def test_missing_license_source_is_not_silently_ignored(tmp_path):
    with pytest.raises(FileNotFoundError):
        release.copy_license_group([], tmp_path / 'licenses', 'Tk')


def test_windows_dlls_license_layout_is_supported(tmp_path):
    home = tmp_path / 'python'
    (home / 'DLLs').mkdir(parents=True)
    (home / 'LICENSE.txt').write_text('Python license')
    (home / 'DLLs' / 'tcllicense.terms').write_text('Tcl license')
    (home / 'DLLs' / 'tklicense.terms').write_text('Tk license')
    package = tmp_path / 'package-license.txt'
    package.write_text('test license')
    release.collect_licenses(tmp_path / 'bundle', python_home=home, package_paths={name: [package] for name in ('tksheet', 'pyinstaller', 'openpyxl', 'et-xmlfile')})
    assert (tmp_path / 'bundle' / 'licenses' / 'Tk' / '01-tklicense.terms').read_text() == 'Tk license'


def test_build_requires_windows_before_running_tools(monkeypatch):
    monkeypatch.setattr(release.sys, 'platform', 'linux')
    with pytest.raises(RuntimeError, match='Windows'):
        release.build()


def test_official_license_url_uses_exact_runtime_tag():
    assert release.official_license_url('Tcl', '8.6.13') == 'https://raw.githubusercontent.com/tcltk/tcl/core-8-6-13/license.terms'
    assert release.official_license_url('Tk', '8.6.13') == 'https://raw.githubusercontent.com/tcltk/tk/core-8-6-13/license.terms'
    with pytest.raises(ValueError):
        release.official_license_url('Tcl', '../main')


def test_missing_runtime_licenses_fetched_and_provenance_recorded(tmp_path):
    home = tmp_path / 'python'
    home.mkdir()
    (home / 'LICENSE.txt').write_text('Python license')
    package = tmp_path / 'package-license.txt'
    package.write_text('package license')
    calls = []
    raw = ('This software is copyrighted. The authors grant permission to use.\n' + 'License fixture text.\n' * 40).encode()
    def download(url):
        calls.append(url)
        return raw
    bundle = tmp_path / 'bundle'
    release.collect_licenses(bundle, python_home=home, package_paths={name: [package] for name in ('tksheet', 'pyinstaller', 'openpyxl', 'et-xmlfile')}, versions={'Tcl': '8.6.13', 'Tk': '8.6.13'}, downloader=download)
    assert len(calls) == 2
    provenance = json.loads((bundle / 'licenses' / 'SOURCES.json').read_text())
    assert provenance[0]['sha256'] == hashlib.sha256(raw).hexdigest()
    assert provenance[0]['version'] == '8.6.13'


def test_invalid_download_does_not_become_license(tmp_path):
    with pytest.raises(ValueError):
        release.download_runtime_license('Tcl', '8.6.13', tmp_path / 'license.terms', lambda url: b'<html>404</html>')
    assert not (tmp_path / 'license.terms').exists()


def test_license_download_failure_stops_preparation(tmp_path):
    def fail(url):
        raise OSError('network unavailable')
    with pytest.raises(OSError):
        release.download_runtime_license('Tcl', '8.6.13', tmp_path / 'license.terms', fail)
    assert not (tmp_path / 'license.terms').exists()


def test_distribution_docs_use_build_version_and_preserve_source(tmp_path, monkeypatch):
    source = tmp_path / 'source'
    source.mkdir()
    for name in release.REQUIRED_DOCS:
        (source / name).write_text('SheetDock {{VERSION}}\nXLSX\n' if name == 'DISTRIBUTION_README.md' else name,
                                   encoding='utf-8')
    monkeypatch.setattr(release, 'ROOT', source)
    bundle = tmp_path / 'bundle'
    bundle.mkdir()
    release.prepare_distribution_docs(bundle, '1.1.0.dev6')
    assert (bundle / 'DISTRIBUTION_README.md').read_text(encoding='utf-8') == 'SheetDock 1.1.0.dev6\nXLSX\n'
    assert '{{VERSION}}' in (source / 'DISTRIBUTION_README.md').read_text(encoding='utf-8')
    assert (bundle / 'LICENSE').read_bytes() == (source / 'LICENSE').read_bytes()


def test_stale_distribution_readme_without_version_token_is_rejected(tmp_path, monkeypatch):
    (tmp_path / 'DISTRIBUTION_README.md').write_text('SheetDock 1.0.0')
    monkeypatch.setattr(release, 'ROOT', tmp_path)
    with pytest.raises(ValueError, match='Version'):
        release.distribution_readme('1.1.0.dev6')


def test_build_pipeline_packages_xlsx_docs_versions_and_preserves_published_zip(tmp_path, monkeypatch):
    source = tmp_path / 'source'
    source.mkdir()
    for name in release.REQUIRED_DOCS:
        (source / name).write_bytes((release.ROOT / name).read_bytes())
    (source / 'pyproject.toml').write_text(f'[project]\nversion = "{__version__}"\n')
    icon = source / 'sheetdock' / 'assets' / 'sheetdock.ico'
    icon.parent.mkdir(parents=True)
    icon.write_bytes(b'icon fixture')
    old_zip = source / 'dist' / 'SheetDock-1.0.0-windows-x64.zip'
    old_zip.parent.mkdir()
    old_zip.write_bytes(b'published v1.0 fixture')
    license_fixture = bundle_fixture(tmp_path) / 'licenses'
    def fake_licenses(target):
        release.shutil.copytree(license_fixture, target / 'licenses')
    commands = []
    def fake_run(command, **kwargs):
        commands.append(command)
        if command[2] == 'PyInstaller':
            bundle = Path(command[command.index('--distpath') + 1]) / 'SheetDock'
            (bundle / '_internal').mkdir(parents=True)
            (bundle / 'SheetDock.exe').write_bytes(b'exe test fixture')
            (bundle / '_internal' / 'python312.dll').write_bytes(b'dll test fixture')
    versions = {'tksheet': '7.6.0', 'pyinstaller': '6.22.3', 'openpyxl': '3.1.5', 'et-xmlfile': '2.0.0'}
    monkeypatch.setattr(release, 'ROOT', source)
    monkeypatch.setattr(release.sys, 'platform', 'win32')
    monkeypatch.setattr(release.struct, 'calcsize', lambda _: 8)
    monkeypatch.setattr(release.platform, 'platform', lambda: 'Windows-test-fixture')
    monkeypatch.setattr(release.importlib.metadata, 'version', versions.__getitem__)
    monkeypatch.setattr(release, 'collect_licenses', fake_licenses)
    monkeypatch.setattr(release.subprocess, 'run', fake_run)
    release.build()
    assert [command[2] for command in commands] == ['pytest', 'PyInstaller']
    output = source / 'dist' / f'SheetDock-{__version__}-windows-x64.zip'
    prefix = f'SheetDock-{__version__}-windows-x64/'
    with ZipFile(output) as archive:
        readme = archive.read(prefix + 'DISTRIBUTION_README.md').decode('utf-8')
        assert f'SheetDock {__version__}' in readme and '{{VERSION}}' not in readme
        assert 'XLSX' in readme and '行列挿入・削除は選択位置から件数を指定' in readme
        info = json.loads(archive.read(prefix + 'BUILD_INFO.json'))
        assert info['sheetdock'] == __version__
        assert info['openpyxl'] == '3.1.5' and info['et-xmlfile'] == '2.0.0'
    assert old_zip.read_bytes() == b'published v1.0 fixture'
    assert output.with_suffix('.zip.sha256').read_text().startswith(hashlib.sha256(output.read_bytes()).hexdigest())
