"""設定JSONの読込と、変更項目だけを反映する保存。"""

import json
import os
import re
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path

from .storage import save_lock
from .widths import fit_widths


@dataclass
class Settings:
    backup_enabled: bool = True
    window_geometry: str | None = None
    recent_files: list[str] = field(default_factory=list)
    dialog_geometry: dict[str, str] = field(default_factory=dict)
    file_widths: dict[str, list[int]] = field(default_factory=dict)


def default_path():
    base = os.environ.get('LOCALAPPDATA')
    return (Path(base) if base else Path.home() / '.local' / 'share') / 'SheetDock' / 'settings.json'


def geometry_parts(value, minimum=(300, 200)):
    if not isinstance(value, str):
        return None
    match = re.fullmatch(r'(\d{1,5})x(\d{1,5})([+-]\d{1,6})([+-]\d{1,6})', value)
    if not match:
        return None
    w, h, x, y = map(int, match.groups())
    return (w, h, x, y) if w >= minimum[0] and h >= minimum[1] else None


def visible_geometry(value, screen_width, screen_height, minimum=(300, 200)):
    parts = geometry_parts(value, minimum)
    if parts is None:
        return None
    w, h, x, y = parts
    w, h = min(w, screen_width), min(h, screen_height)
    x = max(0, min(x, screen_width - w))
    y = max(0, min(y, screen_height - h))
    return f'{w}x{h}+{x}+{y}'


def initial_geometry(saved, screen_width, screen_height):
    restored = visible_geometry(saved, screen_width, screen_height)
    if restored is not None:
        return restored
    width = min(1100, screen_width)
    height = min(700, screen_height)
    return f'{width}x{height}+{(screen_width - width) // 2}+{(screen_height - height) // 2}'


def recent_list(values):
    result, keys = [], set()
    for value in values:
        if not isinstance(value, str) or not value or '\x00' in value:
            continue
        try:
            path = str(Path(value).resolve())
        except (OSError, ValueError):
            continue
        key = os.path.normcase(path)
        if key not in keys:
            result.append(path)
            keys.add(key)
        if len(result) == 10:
            break
    return result


class SettingsStore:
    def __init__(self, path=None):
        self.path = Path(path) if path is not None else default_path()
        self.warning = None

    def load(self):
        self.warning = None
        try:
            value = json.loads(self.path.read_text(encoding='utf-8'))
        except FileNotFoundError:
            return Settings()
        except (OSError, UnicodeError, ValueError):
            self.warning = '設定ファイルを読み込めないため、初期設定で起動しました。'
            return Settings()
        if not isinstance(value, dict) or type(value.get('version')) is not int or value.get('version') != 1:
            self.warning = '設定ファイルの形式が未対応のため、初期設定で起動しました。'
            return Settings()
        backup = value.get('backup_enabled', True)
        geometry = value.get('window_geometry')
        recent = value.get('recent_files', [])
        dialogs = value.get('dialog_geometry', {})
        dialogs = {k: v for k, v in dialogs.items() if isinstance(k, str) and len(k) <= 80 and geometry_parts(v, (80, 50))} if isinstance(dialogs, dict) else {}
        widths = value.get('file_widths', {})
        widths = {k: fit_widths(v, len(v)) for k, v in widths.items() if isinstance(k, str) and isinstance(v, list)} if isinstance(widths, dict) else {}
        return Settings(backup if type(backup) is bool else True, geometry if geometry_parts(geometry) else None, recent_list(recent) if isinstance(recent, list) else [], dict(list(dialogs.items())[-50:]), widths)

    def update(self, *, backup_enabled=None, window_geometry=None, recent_file=None, dialog_key=None, dialog_geometry=None, widths_file=None, column_widths=None):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with save_lock(self.path):
            settings = self.load()
            if backup_enabled is not None:
                if type(backup_enabled) is not bool:
                    raise TypeError('バックアップ設定は有効・無効で指定してください。')
                settings.backup_enabled = backup_enabled
            if window_geometry is not None:
                if geometry_parts(window_geometry) is None:
                    raise ValueError('ウィンドウ位置・サイズが不正です。')
                settings.window_geometry = window_geometry
            if recent_file is not None:
                settings.recent_files = recent_list([str(recent_file), *settings.recent_files])
            if dialog_key is not None:
                if not isinstance(dialog_key, str) or len(dialog_key) > 80 or geometry_parts(dialog_geometry, (80, 50)) is None:
                    raise ValueError('サブ画面の配置が不正です。')
                settings.dialog_geometry[dialog_key] = dialog_geometry
                settings.dialog_geometry = dict(list(settings.dialog_geometry.items())[-50:])
            if widths_file is not None:
                if not isinstance(column_widths, list):
                    raise TypeError('列幅はリストで指定してください。')
                key = os.path.normcase(str(Path(widths_file).resolve()))
                settings.file_widths[key] = fit_widths(column_widths, len(column_widths))
            payload = {'version': 1, **asdict(settings)}
            temporary = None
            try:
                with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=self.path.parent, prefix='.settings.', suffix='.tmp', delete=False) as stream:
                    temporary = Path(stream.name)
                    json.dump(payload, stream, ensure_ascii=False, indent=2)
                    stream.flush()
                    os.fsync(stream.fileno())
                os.replace(temporary, self.path)
            finally:
                if temporary is not None:
                    temporary.unlink(missing_ok=True)
            return settings
