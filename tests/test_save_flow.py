from types import SimpleNamespace

import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat, save_csv
from sheetdock.gui import MultiFileApp, Session
from sheetdock.storage import load_checked
from sheetdock.tabs import Documents


def app_with_external_change(tmp_path):
    path = tmp_path / 'data.csv'
    save_csv(CsvDocument([['ID'], ['00123']], CsvFormat('utf-8', False, '\n', ',')), path)
    doc, stamp = load_checked(path, encoding='utf-8')
    docs = Documents()
    tab = docs.add(path, Session(doc, True), stamp)
    tab.session.editor.set_cell(1, 0, '00999')
    app = MultiFileApp.__new__(MultiFileApp)
    app.root = None
    app.documents = docs
    app.active = tab
    app.session = tab.session
    app.source = tab.source
    app.snapshot = lambda: tab.session.snapshot()
    app.labels = lambda: None
    app.status = SimpleNamespace(set=lambda text: None)
    path.write_bytes(path.read_bytes().replace(b'00123', b'00555'))
    return app, tab


def test_external_change_cancel_keeps_dirty_data(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = app_with_external_change(tmp_path)
    external = tab.source.read_bytes()
    monkeypatch.setattr(gui.messagebox, 'askyesnocancel', lambda *a, **k: None)
    assert not app.save_tab(tab)
    assert tab.source.read_bytes() == external
    assert tab.session.editor.rows[1][0] == '00999'
    assert tab.session.editor.dirty


def test_external_change_overwrite_updates_stamp_and_clean_state(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = app_with_external_change(tmp_path)
    monkeypatch.setattr(gui.messagebox, 'askyesnocancel', lambda *a, **k: True)
    assert app.save_tab(tab)
    assert not tab.session.editor.dirty
    assert load_checked(tab.source, encoding='utf-8')[0].rows[1][0] == '00999'


def test_reload_declined_keeps_edits(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = app_with_external_change(tmp_path)
    monkeypatch.setattr(gui.messagebox, 'askyesnocancel', lambda *a, **k: False)
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: False)
    assert not app.save_tab(tab)
    assert tab.session.editor.rows[1][0] == '00999'
    assert tab.session.editor.dirty


def test_reload_confirmed_replaces_session_and_history(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app, tab = app_with_external_change(tmp_path)
    app.pages = {tab: (None, None)}
    app.render = lambda: None
    monkeypatch.setattr(gui.messagebox, 'askyesnocancel', lambda *a, **k: False)
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: True)
    assert not app.save_tab(tab)
    assert tab.session.editor.rows[1][0] == '00555'
    assert not tab.session.editor.dirty
    assert not tab.session.editor.undo()
    assert app.session is tab.session


@pytest.mark.parametrize('save_as', [False, True])
@pytest.mark.parametrize('encoding,bom', [('utf-8', True), ('utf-8', False), ('cp932', False)])
def test_saved_status_keeps_encoding_bom_and_newline(tmp_path, monkeypatch, save_as, encoding, bom):
    import sheetdock.gui as gui
    path = tmp_path / 'data.csv'
    document = CsvDocument([['名前', 'ID'], ['鈴木', '00123']], CsvFormat(encoding, bom, '\r\n', ','))
    save_csv(document, path)
    loaded, stamp = load_checked(path, encoding=encoding)
    app = MultiFileApp.__new__(MultiFileApp)
    app.documents = Documents()
    tab = app.documents.add(path, Session(loaded, True), stamp)
    tab.session.editor.set_cell(1, 0, '編集')
    app.active = tab
    app.session = tab.session
    app.source = path
    app.root = None
    app.snapshot = lambda: tab.session.snapshot()
    app.labels = lambda: None
    displayed = []
    app.status = SimpleNamespace(set=displayed.append)
    target = tmp_path / 'saved.csv'
    monkeypatch.setattr(gui.filedialog, 'asksaveasfilename', lambda **k: str(target))
    assert (app.save_as_tab(tab) if save_as else app.save_tab(tab))
    status = displayed[-1]
    assert status.startswith('保存完了：')
    assert f' / {encoding} / ' in status
    assert f'BOM={"あり" if bom else "なし"}' in status
    assert 'CRLF' in status
    assert '未保存=なし' in status
    assert str(target if save_as else path) in status


@pytest.mark.parametrize('enabled', [False, True])
def test_gui_backup_setting_controls_file_backup(tmp_path, monkeypatch, enabled):
    import sheetdock.gui as gui
    app, tab = app_with_external_change(tmp_path)
    external = tab.source.read_bytes()
    app.backup_enabled = SimpleNamespace(get=lambda: enabled)
    monkeypatch.setattr(gui.messagebox, 'askyesnocancel', lambda *a, **k: True)
    assert app.save_tab(tab)
    backup = tab.source.with_name(tab.source.name + '.bak')
    assert backup.exists() == enabled
    if enabled:
        assert backup.read_bytes() == external
