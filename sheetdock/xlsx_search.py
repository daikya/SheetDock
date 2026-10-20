"""Occurrence-level XLSX search positions, independent of GUI widgets."""
from dataclasses import dataclass

MAX_MATCHES = 10_000

@dataclass(frozen=True)
class Match:
    sheet: str
    row: int
    column: int
    start: int
    end: int
    value: str


class XlsxSearch:
    def __init__(self):
        self.signature = None
        self.matches = []
        self.index = -1

    def collect(self, session, query, all_sheets, include_formulas):
        matches = []
        for name, row, column, value in session.find(
                query, all_sheets=all_sheets, include_formulas=include_formulas):
            start = 0
            while (position := value.find(query, start)) >= 0:
                end = position + len(query)
                if len(matches) >= MAX_MATCHES:
                    raise ValueError('一致箇所が1万件を超えました。検索文字列または対象シートを絞ってください')
                matches.append(Match(name, row, column, position, end, value))
                start = end
        signature = (id(session), id(session.probe.book), query, all_sheets,
                     include_formulas, None if all_sheets else session.sheet_name,
                     tuple(matches))
        return matches, signature

    def move(self, session, query, all_sheets=True, include_formulas=False, backward=False):
        matches, signature = self.collect(session, query, all_sheets, include_formulas)
        same = signature == self.signature
        self.matches, self.signature = matches, signature
        if not matches:
            self.index = -1
            return None
        self.index = ((self.index + (-1 if backward else 1)) % len(matches)
                      if same else (len(matches) - 1 if backward else 0))
        return matches[self.index]

    def current(self, session, query, all_sheets=True, include_formulas=False):
        matches, signature = self.collect(session, query, all_sheets, include_formulas)
        if signature != self.signature or not 0 <= self.index < len(matches):
            raise ValueError('検索条件またはデータが変わりました。もう一度検索してください')
        return matches[self.index]

    def after_replace(self, session, query, previous, replacement,
                      all_sheets=True, include_formulas=False):
        self.matches, self.signature = self.collect(session, query, all_sheets, include_formulas)
        if not self.matches:
            self.index = -1
            return None
        order = {name: index for index, name in enumerate(session.probe.book.sheetnames)}
        resume = (order[previous.sheet], previous.row, previous.column,
                  previous.start + len(replacement))
        self.index = next((i for i, match in enumerate(self.matches)
            if (order[match.sheet], match.row, match.column, match.start) >= resume), 0)
        return self.matches[self.index]
