import pytest

from sheetdock.gui import MultiFileApp, Session
from sheetdock.storage import load_checked
from sheetdock.tabs import Documents


def make_app(tmp_path, raw, encoding):
    path = tmp_path / 'data.csv'
    path.write_bytes(raw)
    doc, stamp = load_checked(path, encoding=encoding)
    app = MultiFileApp.__new__(MultiFileApp)
    app.documents = Documents()
    app.active = app.documents.add(path, Session(doc, True), stamp)
    app.session = app.active.session
    app.document = doc
    app.root = None
    app.snapshot = lambda: app.session.snapshot()
    app.render = lambda: None
    app.choose_encoding = lambda d: 'utf-8'
    return app


def test_reopen_manual_encoding_replaces_view_after_success(tmp_path):
    app = make_app(tmp_path, '表,ID\n名前,00123\n'.encode('utf-8'), 'cp932')
    assert app.session.format.encoding == 'cp932'
    app.requested_encoding = lambda: 'utf-8'
    app.reopen_encoding()
    assert app.session.format.encoding == 'utf-8'
    assert app.session.editor.rows[0][0] == '表'
    assert app.session is app.active.session
    assert not app.session.editor.dirty


def test_reopen_wrong_encoding_keeps_old_session_and_edits(tmp_path, monkeypatch):
    import sheetdock.gui as gui
    app = make_app(tmp_path, '名前,ID\n鈴木,00123\n'.encode('cp932'), 'cp932')
    app.session.editor.set_cell(1, 0, '未保存')
    original = app.session
    app.requested_encoding = lambda: 'utf-8'
    errors = []
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: True)
    monkeypatch.setattr(gui.messagebox, 'showerror', lambda *a, **k: errors.append(a))
    app.reopen_encoding()
    assert errors
    assert app.session is original
    assert app.session.editor.rows[1][0] == '未保存'
    assert app.session.editor.dirty


def test_reopen_ambiguous_cancel_preserves_session(tmp_path):
    app = make_app(tmp_path, '表,00123\n'.encode('utf-8'), 'utf-8')
    original = app.session
    app.requested_encoding = lambda: 'auto'
    app.choose_encoding = lambda d: None
    app.reopen_encoding()
    assert app.session is original
