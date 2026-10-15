"""Windows: python build_release.py。テスト→exe→ライセンス→配布ZIP。"""

import hashlib
import importlib.metadata
import json
import os
import platform
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import tomllib
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from sheetdock import __version__

ROOT = Path(__file__).resolve().parent
REQUIRED_DOCS = ('LICENSE', 'THIRD_PARTY_NOTICES.md', 'DISTRIBUTION_README.md')


def copy_license_group(paths, destination, label):
    sources = sorted({Path(path).resolve() for path in paths if Path(path).is_file()})
    if not sources:
        raise FileNotFoundError(f'{label}のライセンス原文が見つかりません。ZIP作成を中止します。')
    destination.mkdir(parents=True, exist_ok=True)
    copied = []
    for i, source in enumerate(sources, 1):
        target = destination / f'{i:02d}-{source.name}'
        shutil.copy2(source, target)
        copied.append(target)
    return copied


def package_license_paths(name):
    dist = importlib.metadata.distribution(name)
    return [Path(dist.locate_file(item)) for item in dist.files or [] if Path(str(item)).name.lower().startswith(('license', 'copying', 'copyright'))]


def runtime_versions():
    import tkinter
    root = tkinter.Tk()
    root.withdraw()
    try:
        return {'Tcl': str(root.tk.call('info', 'patchlevel')), 'Tk': str(root.tk.call('package', 'provide', 'Tk'))}
    finally:
        root.destroy()


def official_license_url(name, version):
    if name not in ('Tcl', 'Tk') or re.fullmatch(r'\d+\.\d+\.\d+', version) is None:
        raise ValueError('Tcl/TkのVersionを確定できません。')
    tag = 'core-' + version.replace('.', '-')
    return f'https://raw.githubusercontent.com/tcltk/{name.lower()}/{tag}/license.terms'


def download_runtime_license(name, version, destination, downloader=None):
    url = official_license_url(name, version)
    if downloader is None:
        def downloader(address):
            request = urllib.request.Request(address, headers={'User-Agent': 'SheetDock-release-builder/1.0'})
            with urllib.request.urlopen(request, timeout=30) as response:
                return response.read(128 * 1024)
    data = downloader(url)
    text = data.decode('utf-8')
    if 'copyright' not in text.lower() or 'permission to use' not in text.lower() or len(data) < 500:
        raise ValueError(f'{name}の公式ライセンスの内容を確認できません。')
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(data)
    return {'software': name, 'version': version, 'url': url, 'sha256': hashlib.sha256(data).hexdigest()}


def collect_licenses(bundle, *, python_home=None, package_paths=None, versions=None, downloader=None):
    home = Path(python_home) if python_home is not None else Path(sys.base_prefix)
    output = bundle / 'licenses'
    copied = []
    retrieved = []
    copied += copy_license_group([home / 'LICENSE.txt', home / 'LICENSE'], output / 'Python', 'Python')
    for name, pattern in [('Tcl', 'tcl*'), ('Tk', 'tk*')]:
        sources = [home / f'{name.lower()}license.terms', home / 'DLLs' / f'{name.lower()}license.terms']
        for folder in (home / 'tcl').glob(pattern):
            for prefix in ('license*', 'LICENSE*', 'COPYING*'):
                sources.extend(folder.glob(prefix))
        if not any(Path(path).is_file() for path in sources):
            if versions is None:
                versions = runtime_versions()
            destination = output / name / f'{name.lower()}-{versions[name]}-license.terms'
            retrieved.append(download_runtime_license(name, versions[name], destination, downloader))
            sources = [destination]
        copied += copy_license_group(sources, output / name, name)
    for name in ('tksheet', 'pyinstaller'):
        paths = package_paths[name] if package_paths is not None else package_license_paths(name)
        copied += copy_license_group(paths, output / name, name)
    (output / 'INDEX.txt').write_text('\n'.join(str(p.relative_to(bundle)).replace(os.sep, '/') for p in copied) + '\n', encoding='utf-8')
    (output / 'SOURCES.json').write_text(json.dumps(retrieved, indent=2), encoding='utf-8')


def validate_bundle(bundle):
    if not (bundle / 'SheetDock.exe').is_file() or not (bundle / '_internal').is_dir():
        raise ValueError('SheetDock.exeと_internalがそろっていません。')
    for name in REQUIRED_DOCS:
        if not (bundle / name).is_file():
            raise ValueError(f'配布文書がありません：{name}')
    for group in ('Python', 'Tcl', 'Tk', 'tksheet', 'pyinstaller'):
        if not list((bundle / 'licenses' / group).glob('*')):
            raise ValueError(f'ライセンス文書がありません：{group}')
    if not (bundle / 'licenses' / 'INDEX.txt').is_file() or not (bundle / 'BUILD_INFO.json').is_file():
        raise ValueError('ライセンス一覧またはビルド情報がありません。')
    for path in bundle.rglob('*'):
        parts = {part.lower() for part in path.relative_to(bundle).parts}
        if path.is_symlink() or parts.intersection({'tests', 'probes', '.venv', 'logs', '__pycache__'}):
            raise ValueError(f'配布対象外の項目が混入しています：{path.name}')
        if path.is_file() and (path.name.lower() == 'settings.json' or path.suffix.lower() in {'.csv', '.tsv', '.log', '.bak'}):
            raise ValueError(f'利用者データ・検証データが混入しています：{path.name}')


def create_zip(bundle, output, version):
    validate_bundle(bundle)
    output.parent.mkdir(parents=True, exist_ok=True)
    top = f'SheetDock-{version}-windows-x64'
    descriptor, temporary_name = tempfile.mkstemp(dir=output.parent, suffix='.tmp')
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with ZipFile(temporary, 'w', ZIP_DEFLATED) as archive:
            for path in sorted(bundle.rglob('*')):
                if path.is_file():
                    archive.write(path, f'{top}/{path.relative_to(bundle).as_posix()}')
        with ZipFile(temporary) as archive:
            if archive.testzip() is not None:
                raise ValueError('配布ZIPの整合性検証に失敗しました。')
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    output.with_suffix('.zip.sha256').write_text(f'{digest}  {output.name}\n', encoding='ascii')
    return output


def build():
    if sys.platform != 'win32' or struct.calcsize('P') != 8:
        raise RuntimeError('Windowsの64bit Pythonで実行してください。')
    metadata = tomllib.loads((ROOT / 'pyproject.toml').read_text(encoding='utf-8'))
    if metadata['project']['version'] != __version__:
        raise ValueError('Versionの定義が一致していません。')
    for name, expected in [('tksheet', '7.6.0'), ('pyinstaller', '6.22.3')]:
        if importlib.metadata.version(name) != expected:
            raise ValueError(f'{name}=={expected}をインストールしてください。')
    for name in REQUIRED_DOCS:
        if not (ROOT / name).is_file():
            raise FileNotFoundError(name)
    icon = ROOT / 'sheetdock' / 'assets' / 'sheetdock.ico'
    if not icon.is_file():
        raise FileNotFoundError('アイコンファイルがありません。更新ZIPを展開してください。')
    subprocess.run([sys.executable, '-m', 'pytest', '-q'], cwd=ROOT, check=True)
    # 新しい一時出力でビルドし、既存の配布成果物を失敗時に消さない。
    work = ROOT / 'build'
    work.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(dir=work, prefix='release-') as temporary:
        stage = Path(temporary)
        # ライセンス不足や取得失敗はexeビルド前に検出する。
        notices = stage / 'notices'
        collect_licenses(notices)
        subprocess.run([
            sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed', '--noupx',
            '--name', 'SheetDock', '--paths', str(ROOT),
            '--icon', str(icon), '--add-data', f'{icon.parent}{os.pathsep}sheetdock/assets',
            '--distpath', str(stage / 'dist'), '--workpath', str(stage / 'work'), '--specpath', str(stage),
            '--exclude-module', 'pytest', '--exclude-module', 'sheetdock.benchmark', str(ROOT / 'run_sheetdock.py'),
        ], cwd=ROOT, check=True)
        bundle = stage / 'dist' / 'SheetDock'
        for name in REQUIRED_DOCS:
            shutil.copy2(ROOT / name, bundle / name)
        shutil.copytree(notices / 'licenses', bundle / 'licenses')
        info = {'sheetdock': __version__, 'python': platform.python_version(), 'platform': platform.platform(), 'architecture': 'x64', 'tksheet': importlib.metadata.version('tksheet'), 'pyinstaller': importlib.metadata.version('pyinstaller'), 'built_at_utc': datetime.now(timezone.utc).isoformat(), 'tests': 'passed before build', 'python_tcl_tk_source_changes': 'none'}
        (bundle / 'BUILD_INFO.json').write_text(json.dumps(info, indent=2), encoding='utf-8')
        output = ROOT / 'dist' / f'SheetDock-{__version__}-windows-x64.zip'
        create_zip(bundle, output, __version__)
    print(f'配布ZIP：{output}')
    print('展開してSheetDock.exeを起動し、exe版の実機確認を行ってください。')


if __name__ == '__main__':
    build()
