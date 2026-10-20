import json

import pytest

from sheetdock.settings import SettingsStore, default_path, visible_geometry


def test_default_path_uses_localappdata(tmp_path, monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    assert default_path() == tmp_path / 'SheetDock' / 'settings.json'


def test_missing_settings_defaults_without_creating_directory(tmp_path):
    path = tmp_path / 'new' / 'settings.json'
    settings = SettingsStore(path).load()
    assert settings.backup_enabled
    assert settings.recent_files == []
    assert not path.parent.exists()


def test_save_reload_all_settings(tmp_path):
    path = tmp_path / 'settings.json'
    store = SettingsStore(path)
    store.update(backup_enabled=False, window_geometry='900x600+50+70', recent_file=tmp_path / '日本語.csv')
    loaded = SettingsStore(path).load()
    assert not loaded.backup_enabled
    assert loaded.window_geometry == '900x600+50+70'
    assert loaded.recent_files == [str(tmp_path / '日本語.csv')]


@pytest.mark.parametrize('raw', ['{broken', '[]', '{"version":2}', '{"version":true}', '\ufffd'])
def test_corrupt_or_unsupported_settings_use_defaults(tmp_path, raw):
    path = tmp_path / 'settings.json'
    path.write_text(raw, encoding='utf-8')
    store = SettingsStore(path)
    assert store.load().backup_enabled
    assert store.warning
    assert path.read_text(encoding='utf-8') == raw


def test_invalid_field_types_are_ignored(tmp_path):
    path = tmp_path / 'settings.json'
    path.write_text(json.dumps({'version': 1, 'backup_enabled': 'false', 'window_geometry': 'bad', 'recent_files': [123, None, '', 'bad\x00file']}))
    settings = SettingsStore(path).load()
    assert settings.backup_enabled
    assert settings.window_geometry is None
    assert settings.recent_files == []


def test_recent_files_limited_deduplicated_and_promoted(tmp_path):
    store = SettingsStore(tmp_path / 'settings.json')
    for i in range(12):
        store.update(recent_file=tmp_path / f'{i}.csv')
    assert len(store.load().recent_files) == 10
    assert store.load().recent_files[0] == str(tmp_path / '11.csv')
    store.update(recent_file=tmp_path / '.' / '5.csv')
    settings = store.load()
    assert len(settings.recent_files) == 10
    assert settings.recent_files[0] == str(tmp_path / '5.csv')


def test_two_instances_merge_changed_fields(tmp_path):
    path = tmp_path / 'settings.json'
    a, b = SettingsStore(path), SettingsStore(path)
    a.load(); b.load()
    a.update(backup_enabled=False, recent_file=tmp_path / 'a.csv')
    b.update(window_geometry='800x500+10+20', recent_file=tmp_path / 'b.csv')
    final = a.load()
    assert not final.backup_enabled
    assert final.window_geometry == '800x500+10+20'
    assert final.recent_files == [str(tmp_path / 'b.csv'), str(tmp_path / 'a.csv')]


def test_failed_replace_preserves_previous_settings_and_cleans_temp(tmp_path, monkeypatch):
    import sheetdock.settings as module
    store = SettingsStore(tmp_path / 'settings.json')
    store.update(backup_enabled=True)
    original = store.path.read_bytes()
    def fail(*args):
        raise PermissionError('設定保存失敗')
    monkeypatch.setattr(module.os, 'replace', fail)
    with pytest.raises(PermissionError):
        store.update(backup_enabled=False)
    assert store.path.read_bytes() == original
    assert not list(tmp_path.glob('*.tmp'))
    assert not list(tmp_path.glob('*.sheetdock-save.lock'))


@pytest.mark.parametrize('geometry,expected', [('900x600+5000+5000', '900x600+380+120'), ('900x600-100-100', '900x600+0+0'), ('9999x9999+0+0', '1280x720+0+0'), ('bad', None)])
def test_restore_geometry_stays_on_screen(geometry, expected):
    assert visible_geometry(geometry, 1280, 720) == expected


def test_settings_json_does_not_include_cell_data(tmp_path):
    store = SettingsStore(tmp_path / 'settings.json')
    store.update(recent_file=tmp_path / 'data.csv')
    assert set(json.loads(store.path.read_text())) == {'version', 'backup_enabled', 'window_geometry', 'recent_files', 'dialog_geometry', 'file_widths', 'file_heights'}
