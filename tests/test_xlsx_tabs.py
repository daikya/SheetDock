from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from sheetdock.xlsx_tabs import SheetTabs
from sheetdock.xlsx_gui import XlsxPane, XlsxSession


class Page:
    counter = 0

    def __init__(self, *args, **kwargs):
        Page.counter += 1
        self.key = f'page{Page.counter}'
        self.destroyed = False

    def __str__(self):
        return self.key

    def destroy(self):
        self.destroyed = True


class Book:
    def __init__(self):
        self.selected = ''
        self.labels = {}

    def add(self, page, text):
        self.labels[str(page)] = text

    def forget(self, page):
        self.labels.pop(str(page))

    def select(self, page=None):
        if page is not None:
            self.selected = str(page)
        return self.selected


@pytest.fixture
def tabs(monkeypatch):
    monkeypatch.setattr('sheetdock.xlsx_tabs.ttk.Frame', Page)
    tabs = SheetTabs.__new__(SheetTabs)
    tabs.pages, tabs.current, tabs.book = {}, None, Book()
    changes = []
    tabs.on_select = changes.append
    return tabs, changes


def test_initial_sync_and_queued_events_do_not_trigger_user_change(tabs):
    tabs, changes = tabs
    tabs.sync(['One', 'Two'], 'Two')
    assert list(tabs.book.labels.values()) == ['One', 'Two']
    assert tabs.book.select() == str(tabs.pages['Two'])
    tabs.changed()
    tabs.changed()
    assert changes == []


def test_user_click_and_menu_choice_trigger_once(tabs):
    tabs, changes = tabs
    tabs.sync(['One', 'Two'], 'One')
    tabs.book.select(tabs.pages['Two'])
    tabs.changed()
    assert changes == ['Two']
    tabs.choose('One')
    tabs.changed()  # Tk later delivers the programmatic selection event.
    assert changes == ['Two', 'One']
    tabs.choose('One')
    tabs.choose('missing')
    assert changes == ['Two', 'One']


def test_rerender_reuses_tabs_and_does_not_reset_selection(tabs):
    tabs, changes = tabs
    tabs.sync(['One', 'Two'], 'Two')
    before = tabs.pages.copy()
    tabs.sync(['One', 'Two'], 'Two')
    assert tabs.pages == before
    tabs.changed()
    assert changes == []


@pytest.mark.parametrize('names,current', [
    (['One', 'Two', 'Added'], 'Added'),
    (['Renamed', 'Two'], 'Renamed'),
    (['Two'], 'Two'),
    (['Two', 'One'], 'One'),
])
def test_structure_changes_rebuild_labels_and_discard_stale_pages(tabs, names, current):
    tabs, changes = tabs
    tabs.sync(['One', 'Two'], 'One')
    old = list(tabs.pages.values())
    tabs.sync(names, current)
    assert all(page.destroyed for page in old)
    assert list(tabs.book.labels.values()) == names
    tabs.changed()
    assert changes == []
    assert tabs.current == current


def test_switch_requested_by_history_or_search_has_no_callback_loop(tabs):
    tabs, changes = tabs
    tabs.sync(['One', 'Two'], 'One')
    tabs.sync(['One', 'Two'], 'Two')
    tabs.changed()
    assert tabs.current == 'Two' and changes == []


def test_empty_selection_events_are_ignored(tabs):
    tabs, changes = tabs
    tabs.changed()
    assert changes == []


def test_invalid_sync_does_not_destroy_existing_tabs(tabs):
    tabs, changes = tabs
    tabs.sync(['One'], 'One')
    page = tabs.pages['One']
    with pytest.raises(ValueError):
        tabs.sync(['One'], 'missing')
    assert not page.destroyed and tabs.current == 'One'


def test_many_sheet_menu_reaches_last_sheet(tabs):
    tabs, changes = tabs
    names = [f'Sheet {i}' for i in range(60)]
    tabs.sync(names, names[0])
    commands = []
    tabs.menu = SimpleNamespace(delete=lambda *args: commands.clear(),
        add_command=lambda **kwargs: commands.append(kwargs))
    tabs.fill_menu()
    assert len(commands) == 60
    commands[-1]['command']()
    assert changes == [names[-1]]
    tabs.fill_menu()
    assert commands[-1]['label'].startswith('✓')


def test_pane_switch_preserves_dirty_history_and_validates_sheet(tmp_path):
    book = Workbook()
    book.active.title = 'One'
    book.active['A1'] = 'before'
    book.create_sheet('Two')['A1'] = 'other'
    path = tmp_path / 'sheets.xlsx'
    book.save(path)
    pane = XlsxPane.__new__(XlsxPane)
    pane.session = XlsxSession(path)
    pane.session.edit(0, 0, 'after')
    pane.session.resize_columns({0: 220})
    history = len(pane.session.undo_stack)
    renders = []
    pane.render = lambda: renders.append(pane.session.sheet_name)
    pane.switch_sheet('Two')
    pane.switch_sheet('Two')
    pane.switch_sheet('missing')
    assert renders == ['Two']
    assert pane.session.dirty and len(pane.session.undo_stack) == history
    pane.switch_sheet('One')
    assert pane.session.sheet['A1'].value == 'after'
    assert pane.session.column_widths()[0] == 220
