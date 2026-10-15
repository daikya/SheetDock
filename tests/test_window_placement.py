from types import SimpleNamespace

import pytest

from sheetdock.gui import MultiFileApp
from sheetdock.settings import initial_geometry, SettingsStore


@pytest.mark.parametrize('saved,screen,expected', [(None, (1920, 1080), '1100x700+410+190'), ('bad', (1920, 1080), '1100x700+410+190'), ('900x600+120+80', (1920, 1080), '900x600+120+80'), ('900x600+5000+5000', (1280, 720), '900x600+380+120'), (None, (800, 600), '800x600+0+0')])
def test_initial_geometry_centers_or_restores(saved, screen, expected):
    assert initial_geometry(saved, *screen) == expected


def test_restore_waits_for_layout_and_places_before_showing():
    events = []
    app = MultiFileApp.__new__(MultiFileApp)
    app.settings = SimpleNamespace(window_geometry='900x600+120+80')
    app.restoring_geometry = True
    app.root = SimpleNamespace(update_idletasks=lambda: events.append('layout'), winfo_screenwidth=lambda: 1920, winfo_screenheight=lambda: 1080, geometry=lambda value: events.append(value), deiconify=lambda: events.append('show'))
    app.restore_geometry()
    assert events == ['layout', '900x600+120+80', 'show', 'layout', '900x600+120+80']
    assert app.normal_geometry == '900x600+120+80'
    assert not app.restoring_geometry


def test_initial_configure_does_not_overwrite_saved_position():
    app = MultiFileApp.__new__(MultiFileApp)
    app.restoring_geometry = True
    app.normal_geometry = '900x600+120+80'
    app.root = SimpleNamespace(state=lambda: 'normal', geometry=lambda: '1100x700+0+0')
    app.remember_geometry(SimpleNamespace(widget=app.root))
    assert app.normal_geometry == '900x600+120+80'
    app.restoring_geometry = False
    app.remember_geometry(SimpleNamespace(widget=app.root))
    assert app.normal_geometry == '1100x700+0+0'


def test_normal_and_maximized_close_persist_normal_geometry(tmp_path):
    app = MultiFileApp.__new__(MultiFileApp)
    app.settings_store = SettingsStore(tmp_path / 'settings.json')
    app.root = SimpleNamespace(update_idletasks=lambda: None, state=lambda: 'normal', geometry=lambda: '950x650+150+90')
    app.persist_geometry()
    assert app.settings_store.load().window_geometry == '950x650+150+90'
    app.normal_geometry = '950x650+150+90'
    app.root.state = lambda: 'zoomed'
    app.root.geometry = lambda: '1920x1080+0+0'
    app.persist_geometry()
    assert app.settings_store.load().window_geometry == '950x650+150+90'
