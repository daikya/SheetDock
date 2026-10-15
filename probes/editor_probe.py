"""CSV用編集エンジンの技術検証。GUIとUndo履歴を分離する。"""

from copy import deepcopy


def validate(rows):
    if not isinstance(rows, list) or any(not isinstance(row, list) for row in rows):
        raise TypeError('データは行リストのリストで指定してください。')
    if any(not isinstance(cell, str) for row in rows for cell in row):
        raise TypeError('CSVのセルは文字列のみです。')


class Editor:
    def __init__(self, rows):
        validate(rows)
        self._rows = deepcopy(rows)
        self._saved = deepcopy(rows)
        self._undo = []
        self._redo = []

    @property
    def rows(self):
        return deepcopy(self._rows)

    @property
    def dirty(self):
        return self._rows != self._saved

    @property
    def column_count(self):
        return max((len(row) for row in self._rows), default=0)

    def mark_saved(self):
        # 保存が成功した後にだけ呼び出す。
        self._saved = deepcopy(self._rows)

    def replace_rows(self, rows):
        """GUIから受け取った1操作分のデータを検証し、履歴に記録する。"""
        validate(rows)
        return self._commit(deepcopy(rows))

    def _commit(self, edited):
        validate(edited)
        if edited == self._rows:
            return False
        self._undo.append(self._rows)
        self._rows = edited
        self._redo.clear()
        return True

    def undo(self):
        if not self._undo:
            return False
        self._redo.append(self._rows)
        self._rows = self._undo.pop()
        return True

    def redo(self):
        if not self._redo:
            return False
        self._undo.append(self._rows)
        self._rows = self._redo.pop()
        return True

    @staticmethod
    def _index(value, upper, *, end=False):
        if type(value) is not int or value < 0 or value >= upper + int(end):
            raise IndexError('位置が範囲外です。')

    @staticmethod
    def _count(value):
        if type(value) is not int or value <= 0:
            raise ValueError('件数は正の整数です。')

    def set_cell(self, row, column, value):
        return self.paste(row, column, [[value]])

    def paste(self, row, column, values):
        """既存の表の範囲内へ貼付け。1回の貼付けを1回でUndoする。"""
        validate(values)
        if not values:
            return False
        self._index(row, len(self._rows))
        self._index(column, self.column_count)
        if row + len(values) > len(self._rows) or column + max(map(len, values), default=0) > self.column_count:
            raise IndexError('貼付けが表の範囲を超えます。')
        edited = self.rows
        for offset, cells in enumerate(values):
            target = edited[row + offset]
            if cells and len(target) < column + len(cells):
                target.extend([''] * (column + len(cells) - len(target)))
            for delta, value in enumerate(cells):
                target[column + delta] = value
        return self._commit(edited)

    def insert_rows(self, index, count):
        self._index(index, len(self._rows), end=True)
        self._count(count)
        edited = self.rows
        edited[index:index] = [[''] * self.column_count for _ in range(count)]
        return self._commit(edited)

    def delete_rows(self, indices):
        selected = list(indices)
        for index in selected:
            self._index(index, len(self._rows))
        removed = set(selected)
        return self._commit([deepcopy(row) for i, row in enumerate(self._rows) if i not in removed])

    def insert_columns(self, index, count):
        self._index(index, self.column_count, end=True)
        self._count(count)
        if not self._rows:
            raise ValueError('空の表には先に行を追加してください。')
        edited = self.rows
        for row in edited:
            if len(row) < index:
                row.extend([''] * (index - len(row)))
            row[index:index] = [''] * count
        return self._commit(edited)

    def delete_columns(self, indices):
        selected = list(indices)
        for index in selected:
            self._index(index, self.column_count)
        removed = set(selected)
        return self._commit([[cell for i, cell in enumerate(row) if i not in removed] for row in self._rows])

    def find(self, text):
        if not isinstance(text, str) or not text:
            raise ValueError('検索文字列を指定してください。')
        return [(r, c) for r, row in enumerate(self._rows) for c, value in enumerate(row) if text in value]

    def replace_all(self, old, new):
        if not isinstance(old, str) or not old:
            raise ValueError('検索文字列を指定してください。')
        if not isinstance(new, str):
            raise TypeError('置換後の値は文字列です。')
        count = sum(value.count(old) for row in self._rows for value in row)
        self._commit([[value.replace(old, new) for value in row] for row in self._rows])
        return count
