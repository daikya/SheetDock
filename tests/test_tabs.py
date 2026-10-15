from pathlib import Path

import pytest

from sheetdock.csv_io import CsvDocument, CsvFormat, load_csv, save_csv
from sheetdock.gui import Session
from sheetdock.tabs import Documents


def add(documents, path, encoding='utf-8', bom=False):
    session = Session(CsvDocument([['ID'], ['00123']], CsvFormat(encoding, bom, '\r\n', ',')), True)
    return documents.add(path, session)


def test_tabs_have_independent_data_history_dirty_and_format(tmp_path):
    docs = Documents()
    a = add(docs, tmp_path / 'a.csv', bom=True)
    b = add(docs, tmp_path / 'b.csv', encoding='cp932')
    a.session.editor.set_cell(1, 0, '00999')
    assert a.title == 'a.csv *'
    assert b.title == 'b.csv'
    assert b.session.editor.rows[1][0] == '00123'
    assert not b.session.editor.undo()
    b.session.editor.set_cell(1, 0, '00888')
    a.session.editor.undo()
    assert a.session.editor.rows[1][0] == '00123'
    assert b.session.editor.rows[1][0] == '00888'
    assert a.session.format.bom
    assert b.session.format.encoding == 'cp932'


def test_duplicate_resolved_path_reuses_existing_tab(tmp_path):
    docs = Documents()
    a = add(docs, tmp_path / 'a.csv')
    assert add(docs, tmp_path / '.' / 'a.csv') is a
    assert docs.tabs == [a]


@pytest.mark.parametrize('other', [False, True])
def test_save_as_rejects_own_or_other_open_file(tmp_path, other):
    docs = Documents()
    a = add(docs, tmp_path / 'a.csv')
    b = add(docs, tmp_path / 'b.csv')
    with pytest.raises(ValueError):
        docs.check_save_as(a, b.source if other else a.source)
    assert docs.tabs == [a, b]


def test_save_as_updates_only_saved_tab_and_preserves_format(tmp_path):
    docs = Documents()
    a = add(docs, tmp_path / 'a.csv', bom=True)
    b = add(docs, tmp_path / 'b.csv', encoding='cp932')
    a.session.editor.set_cell(1, 0, '00999')
    b.session.editor.set_cell(1, 0, '00888')
    target = tmp_path / 'saved.csv'
    docs.check_save_as(a, target)
    save_csv(a.session.snapshot(), target)
    docs.saved_as(a, target)
    assert a.source == target
    assert a.title == 'saved.csv'
    assert docs.find(target) is a
    assert docs.find(tmp_path / 'a.csv') is None
    assert b.session.editor.dirty
    assert load_csv(target, encoding='utf-8').format.bom
    a.session.editor.undo()
    assert a.session.editor.dirty


@pytest.mark.parametrize('choice,save_result,expected', [('cancel', False, False), ('discard', False, True), ('save', False, False), ('save', True, True)])
def test_close_choices_do_not_remove_tabs_before_confirmation(tmp_path, choice, save_result, expected):
    docs = Documents()
    tab = add(docs, tmp_path / 'a.csv')
    tab.session.editor.set_cell(1, 0, '00999')
    def save(t):
        if save_result:
            t.session.editor.mark_saved()
        return save_result
    assert docs.may_close(tab, lambda t: choice, save) == expected
    assert docs.tabs == [tab]


def test_clean_tab_closes_without_prompt(tmp_path):
    docs = Documents()
    tab = add(docs, tmp_path / 'a.csv')
    def unexpected(t):
        raise AssertionError('保存確認は不要')
    assert docs.may_close(tab, unexpected, unexpected)
    docs.remove(tab)
    assert docs.tabs == []


def test_reported_save_success_without_clean_state_cannot_close(tmp_path):
    docs = Documents()
    tab = add(docs, tmp_path / 'a.csv')
    tab.session.editor.set_cell(1, 0, '変更')
    assert not docs.may_close(tab, lambda t: 'save', lambda t: True)


def test_failed_write_keeps_path_and_dirty_state(tmp_path, monkeypatch):
    import sheetdock.csv_io as csv_io
    docs = Documents()
    tab = add(docs, tmp_path / 'a.csv')
    tab.session.editor.set_cell(1, 0, '変更')
    target = tmp_path / 'saved.csv'
    def fail(*args):
        raise PermissionError('ファイルが使用中')
    monkeypatch.setattr(csv_io.os, 'replace', fail)
    with pytest.raises(PermissionError):
        save_csv(tab.session.snapshot(), target)
        docs.saved_as(tab, target)
    assert tab.source == tmp_path / 'a.csv'
    assert tab.session.editor.dirty
    assert not target.exists()
