"""CSV用編集エンジン。GUIから独立した編集・Undo/Redo。"""

from copy import deepcopy
from .widths import fit_widths, DEFAULT_WIDTH
from .heights import fit_heights


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
        self.revision = 0
        self.column_widths = fit_widths([], self.column_count)
        self._width_undo = []
        self._width_redo = []
        self.row_heights = fit_heights([], len(rows))
        self._height_undo = []
        self._height_redo = []

    def set_column_widths(self, widths):
        self.column_widths = fit_widths(widths, self.column_count)

    def resize_columns(self, widths):
        """表示列幅だけの変更を、セル編集と同じ順序の履歴に記録する。"""
        widths = fit_widths(widths, self.column_count)
        if widths == self.column_widths:
            return False
        self._undo.append(self._rows)
        self._width_undo.append((self.column_widths[:], widths[:]))
        self._height_undo.append(None)
        self._height_redo.clear()
        self._redo.clear()
        self._width_redo.clear()
        self.column_widths = widths
        self.revision += 1
        return True

    def set_row_heights(self, heights):
        self.row_heights = fit_heights(heights, len(self._rows))

    def resize_rows(self, changes):
        heights = self.row_heights[:]
        for row, pixels in changes.items():
            if type(row) is not int or not 0 <= row < len(heights):
                raise ValueError('表示中の行を指定してください')
            value = fit_heights([pixels], 1)[0]
            if value is None:
                raise ValueError('行の高さが不正です')
            heights[row] = value
        if heights == self.row_heights:
            return False
        self._undo.append(self._rows)
        self._width_undo.append(None)
        self._height_undo.append((self.row_heights[:], heights[:]))
        self._redo.clear()
        self._width_redo.clear()
        self._height_redo.clear()
        self.row_heights = heights
        self.revision += 1
        return True

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

    def _commit(self, edited, widths=None, heights=None):
        validate(edited)
        if edited == self._rows:
            return False
        old_heights = self.row_heights[:]
        new_heights = fit_heights(old_heights if heights is None else heights, len(edited))
        self._height_undo.append((old_heights, new_heights) if heights is not None or len(edited) != len(self._rows) else None)
        self._height_redo.clear()
        self.row_heights = new_heights
        old_widths = self.column_widths[:]
        count = max((len(row) for row in edited), default=0)
        new_widths = fit_widths(old_widths if widths is None else widths, count)
        structural = widths is not None or count != self.column_count
        self._width_undo.append((old_widths, new_widths) if structural else None)
        self._width_redo.clear()
        self.column_widths = new_widths
        self._undo.append(self._rows)
        self._rows = edited
        self._redo.clear()
        self.revision += 1
        return True

    def undo(self):
        if not self._undo:
            return False
        self._redo.append(self._rows)
        self._rows = self._undo.pop()
        heights = self._height_undo.pop()
        self._height_redo.append(heights)
        if heights is not None:
            self.row_heights = heights[0][:]
        widths = self._width_undo.pop()
        self._width_redo.append(widths)
        if widths is not None:
            self.column_widths = widths[0][:]
        self.revision += 1
        return True

    def redo(self):
        if not self._redo:
            return False
        self._undo.append(self._rows)
        self._rows = self._redo.pop()
        heights = self._height_redo.pop()
        self._height_undo.append(heights)
        if heights is not None:
            self.row_heights = heights[1][:]
        widths = self._width_redo.pop()
        self._width_undo.append(widths)
        if widths is not None:
            self.column_widths = widths[1][:]
        self.revision += 1
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
        heights = self.row_heights[:]
        heights[index:index] = [None] * count
        return self._commit(edited, heights=heights)

    def delete_rows(self, indices):
        selected = list(indices)
        for index in selected:
            self._index(index, len(self._rows))
        removed = set(selected)
        return self._commit([deepcopy(row) for i, row in enumerate(self._rows) if i not in removed],
                            heights=[h for i, h in enumerate(self.row_heights) if i not in removed])

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
        widths = self.column_widths[:]
        widths[index:index] = [DEFAULT_WIDTH] * count
        return self._commit(edited, widths)

    def delete_columns(self, indices):
        selected = list(indices)
        for index in selected:
            self._index(index, self.column_count)
        removed = set(selected)
        widths = [width for i, width in enumerate(self.column_widths) if i not in removed]
        return self._commit([[cell for i, cell in enumerate(row) if i not in removed] for row in self._rows], widths)

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
