from types import SimpleNamespace

import pytest

from sheetdock.dialogs import placement, dialog_key, show_placed
from sheetdock.settings import SettingsStore


@pytest.mark.parametrize('saved,parent,size,screen,expected', [
    (None, (100, 80, 1100, 700), (500, 200), (1920, 1080), '500x200+400+330'),
    ('500x200+900+400', (100, 80, 1100, 700), (500, 200), (1920, 1080), '500x200+900+400'),
    (None, (1700, 900, 400, 300), (500, 200), (1920, 1080), '500x200+1420+880'),
    ('200x100+20+30', (100, 80, 1100, 700), (500, 200), (1920, 1080), '200x100+20+30'),
    ('bad', (100, 80, 1100, 700), (500, 200), (1920, 1080), '500x200+400+330'),
])
def test_dialog_parent_center_or_saved_position(saved, parent, size, screen, expected):
    assert placement(saved, parent, size, screen) == expected


def test_dialog_settings_merge_without_losing_main_settings(tmp_path):
    store = SettingsStore(tmp_path / 'settings.json')
    store.update(backup_enabled=False, window_geometry='900x600+80+90')
    store.update(dialog_key='search', dialog_geometry='500x200+100+120')
    store.update(dialog_key='encoding', dialog_geometry='700x300+300+200')
    reloaded = store.load()
    assert not reloaded.backup_enabled
    assert reloaded.window_geometry == '900x600+80+90'
    assert reloaded.dialog_geometry == {'search': '500x200+100+120', 'encoding': '700x300+300+200'}


def test_titles_have_stable_keys_without_storing_text():
    assert dialog_key('input', '見出し編集') == dialog_key('input', '見出し編集')
    assert dialog_key('input', '見出し編集') != dialog_key('input', '挿入')
    assert '見出し' not in dialog_key('input', '見出し編集')


def test_placed_dialog_restores_and_saves_on_destroy(tmp_path, monkeypatch):
    import sheetdock.dialogs as dialogs
    store = SettingsStore(tmp_path / 'settings.json')
    store.update(dialog_key='search', dialog_geometry='500x200+900+400')
    monkeypatch.setattr(dialogs, 'SettingsStore', lambda: store)
    callbacks, geometry = {}, ['1x1+0+0']
    calls = []
    def geom(value=None):
        if value is not None:
            geometry[0] = value
            calls.append(value)
        return geometry[0]
    window = SimpleNamespace(update_idletasks=lambda: None, winfo_reqwidth=lambda: 500, winfo_reqheight=lambda: 200, geometry=geom, deiconify=lambda: calls.append('show'), state=lambda: 'normal', bind=lambda name, func, **k: callbacks.__setitem__(name, func))
    parent = SimpleNamespace(winfo_screenwidth=lambda: 1920, winfo_screenheight=lambda: 1080, winfo_ismapped=lambda: True, winfo_x=lambda: 100, winfo_y=lambda: 80, winfo_width=lambda: 1100, winfo_height=lambda: 700)
    show_placed(window, parent, 'search')
    assert calls == ['500x200+900+400', 'show', '500x200+900+400']
    geom('550x250+110+130')
    callbacks['<Configure>'](SimpleNamespace(widget=window))
    callbacks['<Destroy>'](SimpleNamespace(widget=object()))
    assert store.load().dialog_geometry['search'] == '500x200+900+400'
    callbacks['<Destroy>'](SimpleNamespace(widget=window))
    assert store.load().dialog_geometry['search'] == '550x250+110+130'


def test_corrupt_dialog_entries_are_ignored(tmp_path):
    import json
    path = tmp_path / 'settings.json'
    path.write_text(json.dumps({'version': 1, 'dialog_geometry': {'search': 'bad', 'encoding': '700x300+10+20'}}))
    assert SettingsStore(path).load().dialog_geometry == {'encoding': '700x300+10+20'}


def test_dialog_store_rejects_invalid_geometry(tmp_path):
    with pytest.raises(ValueError):
        SettingsStore(tmp_path / 'settings.json').update(dialog_key='search', dialog_geometry='bad')
