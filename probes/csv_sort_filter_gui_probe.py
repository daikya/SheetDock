"""Standalone CSV projection GUI. Production SheetDock does not import this probe."""
from __future__ import annotations

import csv
import io
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from tksheet import Sheet

from sheetdock.csv_io import load_csv
if __package__:
    from .csv_sort_filter_probe import SortFilterEditor, save_copy
else:
    from csv_sort_filter_probe import SortFilterEditor, save_copy


def history_shortcut(event, window_system):
    if not event.state & 0x0004:
        return None
    # Windows IME can report a non-Latin keysym; use its physical Z key too.
    if event.keysym.lower() == 'z' or (window_system == 'win32' and event.keycode == 90):
        return bool(event.state & 0x0001)
    return None


class ProjectionSession:
    def __init__(self, document, *, header=True):
        self.editor = SortFilterEditor(document.rows, header=header)
        self.format = document.format
        self.mapping = []
        self.view_token = None

    def token(self):
        return self.editor.revision, self.editor.header, self.editor.condition

    def view(self):
        rows = self.editor.rows
        offset = int(self.editor.header and bool(rows))
        self.mapping = [i for i in self.editor.visible_indices() if i >= offset]
        self.view_token = self.token()
        return (rows[0] if offset else None,
                [rows[i] for i in self.mapping], self.mapping[:], self.view_token)

    def data_row(self, row, token):
        if token != self.view_token or token != self.token():
            raise ValueError('表示条件やデータが変わりました。セルを選び直してください')
        if type(row) is not int or not 0 <= row < len(self.mapping):
            raise IndexError('表示中の本文セルを選択してください')
        return self.mapping[row]

    def paste(self, row, column, values, token):
        self.data_row(row, token)
        self.editor.check_column(column)
        if (not isinstance(values, list) or not values or
                any(not isinstance(cells, list) or not cells for cells in values) or
                len({len(cells) for cells in values}) != 1 or
                any(not isinstance(value, str) for cells in values for value in cells)):
            raise ValueError('空でない長方形の文字列範囲を指定してください')
        if row + len(values) > len(self.mapping) or column + len(values[0]) > self.editor.column_count:
            raise IndexError('貼り付けが表示範囲を超えます。全体を中止しました')
        rows = self.editor.rows
        for offset, cells in enumerate(values):
            target = self.mapping[row + offset]
            width = column + len(cells)
            rows[target].extend([''] * max(0, width - len(rows[target])))
            rows[target][column:width] = cells
        return self.editor.replace_rows(rows)

    def edit(self, row, column, value, token):
        return self.paste(row, column, [[value]], token)

    def save_full_copy(self, path):
        # Projection is never used as the save source, including zero matches.
        full = SortFilterEditor(self.editor.rows, header=self.editor.header)
        save_copy(full, self.format, path)
        self.editor.mark_saved()

    @property
    def history_counts(self):
        return len(self.editor._undo), len(self.editor._redo)

    def history(self, redo=False):
        return self.editor.redo() if redo else self.editor.undo()


class ProjectionGui:
    def __init__(self, root):
        self.root, self.session, self.source = root, None, None
        self.mapping, self.view_token = [], None
        root.title('SheetDock — CSV並べ替え・フィルター 独立GUI検証（履歴修正2）')
        root.geometry('1060x680')
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)
        self.encoding = tk.StringVar(value='utf-8')
        self.delimiter = tk.StringVar(value='カンマ')
        self.header = tk.BooleanVar(value=True)
        self.column = tk.StringVar(value='1')
        self.mode = tk.StringVar(value='文字列')
        self.descending = tk.BooleanVar(value=False)
        self.operator = tk.StringVar(value='部分一致')
        self.query = tk.StringVar()
        self.status = tk.StringVar(value='検証用CSVを開いてください。元ファイルは上書きしません。')
        self.detail = tk.StringVar()
        self.history_notice = tk.StringVar(value='履歴修正2：セルを変更するとUndo件数が増えます。')
        self.key_notice = tk.StringVar(value='履歴キー：未入力')
        opening = ttk.Frame(root)
        opening.grid(row=0, column=0, sticky='ew', padx=8, pady=6)
        ttk.Label(opening, text='次に開くCSV').pack(side='left')
        ttk.Combobox(opening, textvariable=self.encoding, values=('utf-8', 'cp932'), state='readonly', width=8).pack(side='left', padx=4)
        ttk.Combobox(opening, textvariable=self.delimiter, values=('カンマ', 'タブ', 'セミコロン'), state='readonly', width=10).pack(side='left')
        ttk.Checkbutton(opening, text='見出しあり', variable=self.header).pack(side='left', padx=5)
        ttk.Button(opening, text='開く', command=self.open).pack(side='left')
        actions = ttk.Frame(root)
        actions.grid(row=1, column=0, sticky='ew', padx=8, pady=3)
        for label, command in [('全行を別名保存', self.save), ('セル編集', self.edit),
                               ('範囲貼り付け', self.paste)]:
            ttk.Button(actions, text=label, command=command).pack(side='left', padx=3)
        self.undo_button = ttk.Button(actions, text='Undo', command=lambda: self.history(False))
        self.undo_button.pack(side='left', padx=3)
        self.redo_button = ttk.Button(actions, text='Redo', command=lambda: self.history(True))
        self.redo_button.pack(side='left', padx=3)
        ttk.Label(actions, textvariable=self.history_notice).pack(side='left', padx=8)
        controls = ttk.Frame(root)
        controls.grid(row=2, column=0, sticky='ew', padx=8, pady=4)
        ttk.Label(controls, text='対象列番号').pack(side='left')
        ttk.Spinbox(controls, from_=1, to=1000, textvariable=self.column, width=5).pack(side='left', padx=4)
        ttk.Combobox(controls, textvariable=self.mode, values=('文字列', '数値'), state='readonly', width=7).pack(side='left')
        ttk.Checkbutton(controls, text='降順', variable=self.descending).pack(side='left')
        ttk.Button(controls, text='並べ替え', command=self.sort).pack(side='left', padx=4)
        ttk.Combobox(controls, textvariable=self.operator, values=('部分一致', '完全一致', '空', '空以外'), state='readonly', width=9).pack(side='left')
        ttk.Entry(controls, textvariable=self.query, width=16).pack(side='left', padx=4)
        ttk.Button(controls, text='フィルター適用', command=self.filter).pack(side='left')
        ttk.Button(controls, text='解除', command=self.clear).pack(side='left', padx=4)
        ttk.Label(root, text='行番号は見出し込みのデータ行番号。編集はセル選択→「セル編集」。\n'
                  'フィルター中も非表示行を含む全行を保存します。並べ替え前にはフィルターを解除してください。',
                  wraplength=1020).grid(row=3, column=0, sticky='ew', padx=8, pady=4)
        self.sheet = Sheet(root, data=[], headers=[], cell_auto_resize_enabled=False)
        self.sheet.grid(row=4, column=0, sticky='nsew', padx=8, pady=4)
        # Native mutation/structure/history bindings are intentionally excluded.
        self.sheet.enable_bindings('single_select', 'drag_select', 'row_select', 'column_select',
                                   'select_all', 'arrowkeys', 'copy')
        self.sheet.bind('<<SheetSelect>>', self.selection_detail)
        ttk.Label(root, textvariable=self.status, wraplength=1020).grid(row=5, column=0, sticky='ew', padx=8, pady=4)
        ttk.Label(root, textvariable=self.detail).grid(row=6, column=0, sticky='ew', padx=8, pady=4)
        ttk.Label(root, textvariable=self.key_notice, wraplength=1020).grid(row=7, column=0, sticky='ew', padx=8, pady=3)
        # A first bindtag handles physical keys before widget/class/virtual-event
        # bindings can consume them. It is scoped to this probe window only.
        self.history_tag = f'ProjectionHistory_{id(self)}'
        root.bind_class(self.history_tag, '<KeyPress>', self.route_history_key)
        self.install_history_tag(root)
        for key, callback in [('<Control-v>', self.paste), ('<Control-V>', self.paste)]:
            root.bind(key, lambda event, callback=callback: self.table_key(event, callback))
        root.protocol('WM_DELETE_WINDOW', self.close)

    def install_history_tag(self, widget):
        tags = widget.bindtags()
        widget.bindtags((self.history_tag,) + tuple(tag for tag in tags if tag != self.history_tag))
        for child in widget.winfo_children():
            self.install_history_tag(child)

    def route_history_key(self, event):
        redo = history_shortcut(event, self.root.tk.call('tk', 'windowingsystem'))
        if redo is None:
            return None
        widget_class = event.widget.winfo_class()
        self.key_notice.set(f'履歴キー受信: {event.keysym} / keycode={event.keycode} / '
                            f'state={event.state} / {widget_class}')
        if widget_class in {'Entry', 'TEntry', 'Text', 'TCombobox', 'Spinbox', 'TSpinbox'}:
            self.key_notice.set(self.key_notice.get() + ' / 入力欄のため表の履歴は操作しません')
            return None
        self.history(redo)
        return 'break'

    def table_key(self, event, callback):
        if self.root.focus_get() not in (self.sheet.MT, self.sheet.CH, self.sheet.RI):
            return None
        callback()
        return 'break'

    def run(self, action):
        if not self.session:
            return
        try:
            action()
            self.render()
            self.sheet.MT.focus_set()
        except (ValueError, TypeError, IndexError, OSError, tk.TclError, csv.Error) as error:
            messagebox.showerror('操作中止', str(error), parent=self.root)

    def may_discard(self):
        return not (self.session and self.session.editor.dirty) or messagebox.askyesno(
            '未保存の編集', '未保存の編集を破棄しますか？', parent=self.root)

    def open(self):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[('CSV / TSV', '*.csv *.tsv')])
        if path:
            self.open_path(path)

    def open_path(self, path):
        try:
            doc = load_csv(path, encoding=self.encoding.get(), delimiter={'カンマ': ',', 'タブ': '\t', 'セミコロン': ';'}[self.delimiter.get()])
            if sum(map(len, doc.rows)) > 100_000:
                raise ValueError('この独立GUIは10万セル以内の検証用CSVを対象とします')
            session = ProjectionSession(doc, header=self.header.get())
            if not self.may_discard():
                return
            self.session, self.source = session, Path(path)
            self.render()
            self.sheet.MT.focus_set()
        except (ValueError, TypeError, OSError) as error:
            messagebox.showerror('読込中止', str(error), parent=self.root)

    def render(self):
        header, rows, self.mapping, self.view_token = self.session.view()
        editor = self.session.editor
        count = editor.column_count
        padded = [row + [''] * (count - len(row)) for row in rows]
        self.sheet.set_sheet_data(padded, delete_options=True)
        self.sheet.headers([f'{i + 1}: {header[i] if header is not None and i < len(header) else ""}' for i in range(count)])
        self.sheet.row_index([str(index + 1) for index in self.mapping])
        self.sheet.deselect('all')
        self.sheet.reset_undos()
        total = len(editor.rows) - int(editor.header and bool(editor.rows))
        fmt = self.session.format
        self.status.set(f'{self.source.name} / 表示 {len(rows)} / 全本文 {total}行 / '
                        f'フィルター: {"あり" if editor.condition else "なし"} / '
                        f'未保存: {"あり" if editor.dirty else "なし"} / '
                        f'{fmt.encoding} / BOM={fmt.bom} / 全行保存 / '
                        f'Undo {self.session.history_counts[0]}件 / Redo {self.session.history_counts[1]}件')
        undo_count, redo_count = self.session.history_counts
        self.undo_button.configure(state='normal' if undo_count else 'disabled')
        self.redo_button.configure(state='normal' if redo_count else 'disabled')
        self.detail.set('セルを選択してください。見出しの編集・行列の追加削除は今回対象外です。')

    def selected(self):
        selected = self.sheet.get_currently_selected()
        if not selected or selected.row >= len(self.mapping):
            raise ValueError('表示中の本文セルを選択してください')
        return selected.row, selected.column, self.view_token

    def selection_detail(self, event=None):
        if not self.session:
            return
        try:
            row, column, token = self.selected()
            target = self.session.data_row(row, token)
            self.detail.set(f'表示本文{row + 1}行 → データ{target + 1}行 / {column + 1}列')
        except (ValueError, IndexError):
            self.detail.set('表示中の本文セルを選択してください')

    def edit(self):
        def action():
            session = self.session
            row, column, token = self.selected()
            target = session.data_row(row, token)
            current = session.editor.value(session.editor.rows[target], column)
            value = simpledialog.askstring('セル編集', f'データ{target + 1}行・{column + 1}列', initialvalue=current, parent=self.root)
            if value is not None:
                if session is not self.session:
                    raise ValueError('編集中に対象ファイルが変わりました')
                session.edit(row, column, value, token)
        self.run(action)

    def paste(self):
        def action():
            row, column, token = self.selected()
            values = list(csv.reader(io.StringIO(self.root.clipboard_get(), newline=''), delimiter='\t', strict=True))
            self.session.paste(row, column, values, token)
        self.run(action)

    def sort(self):
        self.run(lambda: self.session.editor.sort_rows(int(self.column.get()) - 1,
                 mode={'文字列': 'text', '数値': 'number'}[self.mode.get()], descending=self.descending.get()))

    def filter(self):
        self.run(lambda: self.session.editor.set_filter(int(self.column.get()) - 1,
                 operator={'部分一致': 'contains', '完全一致': 'equals', '空': 'empty', '空以外': 'nonempty'}[self.operator.get()], query=self.query.get()))

    def clear(self):
        self.run(lambda: self.session.editor.clear_filter())

    def history(self, redo=False):
        if not self.session:
            self.history_notice.set('ファイルを開いてください')
            return
        label = 'Redo' if redo else 'Undo'
        try:
            changed = self.session.history(redo)
            self.render()
            self.sheet.MT.focus_set()
            self.history_notice.set(f'{label}: {"1操作を反映しました" if changed else "対象の履歴はありません"}')
        except (ValueError, TypeError, IndexError, tk.TclError) as error:
            self.history_notice.set(f'{label}: エラー')
            messagebox.showerror('履歴操作中止', str(error), parent=self.root)

    def save(self):
        if not self.session:
            return
        path = filedialog.asksaveasfilename(parent=self.root, initialdir=str(self.source.parent),
                   initialfile=self.source.stem + '_gui_copy.csv', defaultextension='.csv',
                   filetypes=[('CSV', '*.csv')])
        if path:
            def action():
                self.session.save_full_copy(path)
                messagebox.showinfo('保存完了', '非表示行を含む全行を別名保存し、再読込比較しました。', parent=self.root)
            self.run(action)

    def close(self):
        if self.may_discard():
            self.root.destroy()


def main():
    root = tk.Tk()
    ProjectionGui(root)
    root.mainloop()


if __name__ == '__main__':
    main()
