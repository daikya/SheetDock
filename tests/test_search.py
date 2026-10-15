import pytest

from sheetdock.editor import Editor
from sheetdock.search import SearchState, find_matches


def test_multiple_occurrences_and_cell_order():
    found = find_matches([['日本語日本語', '日本語'], [], ['日本語']], '日本語')
    assert [m.position for m in found] == [(0, 0, 0), (0, 0, 3), (0, 1, 0), (2, 0, 0)]


def test_next_and_previous_wrap_in_both_directions():
    e = Editor([['aa', 'a']])
    s = SearchState()
    assert [s.navigate(e, 'a').position for _ in range(4)] == [(0, 0, 0), (0, 0, 1), (0, 1, 0), (0, 0, 0)]
    assert s.navigate(e, 'a', backwards=True).position == (0, 1, 0)
    assert s.navigate(e, 'a', backwards=True).position == (0, 0, 1)


def test_first_previous_selects_last_match():
    e = Editor([['aa', 'a']])
    assert SearchState().navigate(e, 'a', backwards=True).position == (0, 1, 0)


def test_query_change_resets_cursor_and_case_is_sensitive():
    e = Editor([['Aaa', 'B']])
    s = SearchState()
    s.navigate(e, 'a')
    s.navigate(e, 'a')
    assert s.navigate(e, 'A').position == (0, 0, 0)
    assert s.navigate(e, 'b') is None


@pytest.mark.parametrize('replacement', ['', '置換', '日本語追加'])
def test_replace_one_undo_redo_and_next_skips_inserted_text(replacement):
    e = Editor([['日本語日本語', '00123']])
    s = SearchState()
    s.navigate(e, '日本語')
    assert s.replace_current(e, '日本語', replacement)
    assert e.rows == [[replacement + '日本語', '00123']]
    assert s.navigate(e, '日本語').start == len(replacement)
    assert e.undo()
    assert e.rows == [['日本語日本語', '00123']]
    assert e.redo()
    assert e.rows == [[replacement + '日本語', '00123']]


def test_replace_without_search_selects_first():
    e = Editor([['aa']])
    assert SearchState().replace_current(e, 'a', 'b')
    assert e.rows == [['ba']]


def test_no_match_does_not_create_history():
    e = Editor([['00123']])
    assert not SearchState().replace_current(e, 'x', 'y')
    assert not e.undo()


@pytest.mark.parametrize('change', [lambda e: e.set_cell(0, 0, 'a'), lambda e: e.insert_rows(0, 1), lambda e: e.delete_columns([0])])
def test_stale_match_after_edit_or_structure_change_is_rejected(change):
    e = Editor([['aa', 'aa']])
    s = SearchState()
    s.navigate(e, 'a')
    change(e)
    expected = e.rows
    with pytest.raises(ValueError, match='変更'):
        s.replace_current(e, 'a', 'b')
    assert e.rows == expected


def test_undo_also_invalidates_selected_match():
    e = Editor([['aa']])
    e.set_cell(0, 0, 'aaa')
    s = SearchState()
    s.navigate(e, 'a')
    e.undo()
    with pytest.raises(ValueError):
        s.replace_current(e, 'a', 'b')


def test_states_are_independent_between_tabs():
    e = Editor([['aa']])
    a, b = SearchState(), SearchState()
    a.navigate(e, 'a')
    a.navigate(e, 'a')
    assert b.navigate(e, 'a').start == 0
    assert a.current.start == 1


def test_replace_all_is_one_undo_and_search_recomputes():
    e = Editor([['aa'], ['aa']])
    s = SearchState()
    s.navigate(e, 'a')
    assert e.replace_all('a', 'b') == 4
    assert s.navigate(e, 'a') is None
    assert e.undo()
    assert e.rows == [['aa'], ['aa']]


def test_overlapping_matches_use_nonoverlapping_replace_semantics():
    assert [m.start for m in find_matches([['aaa']], 'aa')] == [0]


@pytest.mark.parametrize('text', ['', None])
def test_empty_or_invalid_query_rejected(text):
    with pytest.raises(ValueError):
        find_matches([['value']], text)
