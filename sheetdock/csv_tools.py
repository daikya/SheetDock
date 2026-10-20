"""CSV row sorting and non-destructive filter conditions."""
from decimal import Decimal, DecimalException
import re

from .editor import Editor, validate


class SortFilterEditor(Editor):
    def __init__(self, rows, *, header=True):
        super().__init__(rows)
        if type(header) is not bool:
            raise TypeError('見出し指定はboolです')
        self.header = header
        self.condition = None

    def check_column(self, column):
        self._index(column, self.column_count)

    @staticmethod
    def value(row, column):
        return row[column] if column < len(row) else ''

    def sort_rows(self, column, *, mode='text', descending=False):
        self.check_column(column)
        if mode not in ('text', 'number') or type(descending) is not bool:
            raise ValueError('並べ替え条件が不正です')
        if self.condition is not None:
            raise ValueError('並べ替え前にフィルターを解除してください')
        rows = self.rows
        start = int(self.header and bool(rows))
        filled, blanks = [], []
        for index in range(start, len(rows)):
            value = self.value(rows[index], column)
            if value == '':
                blanks.append(index)
                continue
            key = value
            if mode == 'number':
                # Explicit numeric interpretation only; original CSV strings stay intact.
                if not re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?', value.strip()):
                    raise ValueError(f'{index + 1}行目を数値として比較できません')
                try:
                    key = Decimal(value.strip())
                except DecimalException as error:
                    raise ValueError(f'{index + 1}行目の数値が範囲外です') from error
            filled.append((key, index))
        filled.sort(key=lambda item: item[0], reverse=descending)
        order = list(range(start)) + [index for _, index in filled] + blanks
        return self._commit([rows[index] for index in order],
                            heights=[self.row_heights[index] for index in order])

    def set_filter(self, column, *, operator='contains', query=''):
        self.check_column(column)
        if operator not in ('contains', 'equals', 'empty', 'nonempty') or not isinstance(query, str):
            raise ValueError('フィルター条件が不正です')
        self.condition = (column, operator, query)

    def clear_filter(self):
        self.condition = None

    def visible_indices(self):
        rows = self.rows
        if self.condition is None:
            return list(range(len(rows)))
        column, operator, query = self.condition
        visible = []
        for index, row in enumerate(rows):
            value = self.value(row, column)
            matches = {'contains': lambda: query in value,
                       'equals': lambda: query == value,
                       'empty': lambda: value == '',
                       'nonempty': lambda: value != ''}[operator]()
            if (self.header and index == 0) or matches:
                visible.append(index)
        return visible

    def paste_visible(self, row, column, values):
        """One atomic paste across the currently visible body rows only."""
        validate(values)
        if not values or not values[0] or len({len(cells) for cells in values}) != 1:
            raise ValueError('空でない長方形の貼り付けを指定してください')
        indices = self.visible_indices()
        if self.header and self._rows:
            indices = [index for index in indices if index != 0]
        self._index(row, len(indices))
        self.check_column(column)
        if row + len(values) > len(indices) or column + len(values[0]) > self.column_count:
            raise IndexError('貼り付けが表示範囲を超えます。全体を中止しました')
        edited = self.rows
        for offset, cells in enumerate(values):
            target = edited[indices[row + offset]]
            end = column + len(cells)
            target.extend([''] * max(0, end - len(target)))
            target[column:end] = cells
        return self._commit(edited)
