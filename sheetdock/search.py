"""文字列の一致箇所を順に検索し、現在の1箇所だけ置換する。"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Match:
    row: int
    column: int
    start: int
    end: int
    value: str

    @property
    def position(self):
        return self.row, self.column, self.start


def find_matches(rows, text):
    if not isinstance(text, str) or not text:
        raise ValueError('検索文字列を指定してください。')
    result = []
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            start = 0
            while (index := value.find(text, start)) != -1:
                result.append(Match(r, c, index, index + len(text), value))
                start = index + len(text)
    return result


class SearchState:
    def __init__(self):
        self.query = None
        self.anchor = None
        self.current = None
        self.revision = None

    def reset(self):
        self.query = self.anchor = self.current = None
        self.revision = None

    def navigate(self, editor, text, *, backwards=False):
        matches = find_matches(editor.rows, text)
        if text != self.query:
            self.anchor = None
        self.query = text
        self.revision = editor.revision
        if not matches:
            self.anchor = self.current = None
            return None
        if self.anchor is None:
            selected = matches[-1] if backwards else matches[0]
        elif backwards:
            selected = next((m for m in reversed(matches) if m.position < self.anchor), matches[-1])
        else:
            selected = next((m for m in matches if m.position > self.anchor), matches[0])
        self.current = selected
        self.anchor = selected.position
        return selected

    def replace_current(self, editor, text, replacement):
        if not isinstance(replacement, str):
            raise TypeError('置換後の値は文字列です。')
        if self.current is None or self.query != text:
            self.navigate(editor, text)
        match = self.current
        if match is None:
            return False
        if self.revision != editor.revision:
            self.reset()
            raise ValueError('検索後に表が変更されました。もう一度検索してください。')
        rows = editor.rows
        if match.row >= len(rows) or match.column >= len(rows[match.row]) or rows[match.row][match.column] != match.value:
            self.reset()
            raise ValueError('検索後にセルが変更されました。もう一度検索してください。')
        changed = editor.set_cell(match.row, match.column, match.value[:match.start] + replacement + match.value[match.end:])
        # 次へ進む際、今挿入した文字列内の一致はスキップする。
        self.anchor = (match.row, match.column, match.start + len(replacement) - 1)
        self.current = None
        return changed
