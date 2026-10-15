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
    for group in ('Python', 'Tcl', 'Tk', 'tksheet', 'pyinstaller'):
        folder = bundle / 'licenses' / group
        folder.mkdir(parents=True)
        (folder / 'LICENSE.txt').write_text('test license fixture')
    (bundle / 'licenses' / 'INDEX.txt').write_text('test index')
    return bundle


def test_version_and_mit_license_consistency():
    root = Path(__file__).resolve().parents[1]
    metadata = tomllib.loads((root / 'pyproject.toml').read_text())
    assert metadata['project']['version'] == __version__ == '1.0.0'
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


@pytest.mark.parametrize('relative', ['settings.json', 'private.csv', 'private.tsv', 'logs/a.log', 'tests/test_private.py'])
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
    for name in ('tksheet', 'pyinstaller'):
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
    release.collect_licenses(tmp_path / 'bundle', python_home=home, package_paths={'tksheet': [package], 'pyinstaller': [package]})
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
    release.collect_licenses(bundle, python_home=home, package_paths={'tksheet': [package], 'pyinstaller': [package]}, versions={'Tcl': '8.6.13', 'Tk': '8.6.13'}, downloader=download)
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
