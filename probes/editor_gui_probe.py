"""実行: python probes/editor_gui_probe.py。編集エンジンとGUIの接続検証。"""

import copy
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from csv_gui_probe import CsvGuiProbe, split_rows
from csv_probe import CsvDocument
from editor_probe import Editor


class Session:
    def __init__(self, document, use_header):
        self.format = document.format
        self.has_header = use_header
        self.editor = Editor(document.rows)

    def view(self):
        return split_rows(self.editor.rows, self.has_header)

    @property
    def offset(self):
        return int(self.has_header and bool(self.editor.rows))

    def accept_body(self, body):
        header, _ = self.view()
        rows = copy.deepcopy(body)
        if header is not None:
            rows.insert(0, header)
        return self.editor.replace_rows(rows)

    def snapshot(self):
        return CsvDocument(self.editor.rows, self.format)


class EditorGuiProbe(CsvGuiProbe):
    def __init__(self, root):
        self.session = None
        self.rendering = False
        super().__init__(root)
        root.title('SheetDock — 編集エンジン接続検証')
        self.sheet.disable_bindings('undo', 'redo', 'find', 'replace')
        self.sheet.bind('<<SheetModified>>', self.modified)
        # 既存ボタンもキー操作もエンジンの履歴へ向ける。
        for frame in root.winfo_children():
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Button):
                    text = widget.cget('text')
                    if text == 'Undo': widget.configure(command=lambda: self.history(False))
                    elif text == 'Redo': widget.configure(command=lambda: self.history(True))
                    elif text == '検索・置換': widget.configure(text='検索', command=self.find)
        root.bind('<Control-z>', lambda event: self.history(False))
        root.bind('<Control-Shift-Z>', lambda event: self.history(True))
        root.bind('<Control-Shift-z>', lambda event: self.history(True))
        toolbar = ttk.Frame(root)
        toolbar.pack(fill='x', padx=10, pady=5)
        for label, action in [('行を挿入', lambda: self.insert(False)), ('選択行を削除', lambda: self.delete(False)), ('列を挿入', lambda: self.insert(True)), ('選択列を削除', lambda: self.delete(True)), ('見出し編集', self.edit_header), ('一括置換', self.replace)]:
            ttk.Button(toolbar, text=label, command=action).pack(side='left', padx=3)
        # 基底クラスの表示専用見出しに関する案内を差し替える。
        for widget in root.winfo_children():
            if isinstance(widget, ttk.Label) and str(widget.cget('text')).startswith('見出しは表示専用'):
                widget.configure(text='行・列の番号で複数選択し、ボタンで操作します。見出し編集もUndo可能です。')

    def open(self):
        previous = self.document
        super().open()
        if self.document is not None and self.document is not previous:
            self.session = Session(self.document, self.header is not None)
            self.render()

    def render(self):
        if self.session is None:
            return
        self.rendering = True
        try:
            header, body = self.session.view()
            self.header = header
            self.sheet.set_sheet_data(body, delete_options=True)
            self.sheet.headers(header if header is not None else [])
            self.sheet.reset_undos()
            self.status.set(f'{self.source} / 未保存={self.session.editor.dirty} / 履歴は編集エンジンで管理')
        finally:
            self.rendering = False

    def modified(self, event=None):
        if self.rendering or self.session is None:
            return
        try:
            self.session.accept_body(self.sheet.get_sheet_data())
            # tksheet内部の履歴は保持しない。
            self.sheet.reset_undos()
            self.status.set(f'{self.source} / 未保存={self.session.editor.dirty}')
        except Exception as exc:
            self.render()
            messagebox.showerror('編集失敗', str(exc), parent=self.root)

    def snapshot(self):
        if self.session is None:
            return super().snapshot()
        self.sheet.close_text_editor(set_data=True)
        self.modified()
        return self.session.snapshot()

    def save(self):
        super().save()
        if self.session and self.last_saved_rows == self.session.editor.rows:
            self.session.editor.mark_saved()

    def history(self, redo):
        if self.session:
            self.snapshot()
            (self.session.editor.redo if redo else self.session.editor.undo)()
            self.render()
        return 'break'

    def act(self, operation):
        if self.session is None:
            return
        try:
            self.snapshot()
            operation(self.session.editor)
            self.render()
        except Exception as exc:
            messagebox.showerror('操作失敗', str(exc), parent=self.root)

    def selected(self, columns):
        return sorted(self.sheet.get_selected_columns() if columns else self.sheet.get_selected_rows())

    def insert(self, columns):
        if self.session is None:
            return
        selected = self.selected(columns)
        count = simpledialog.askinteger('挿入', '挿入する列数' if columns else '挿入する行数', minvalue=1, maxvalue=1000, parent=self.root)
        if count is None:
            return
        if columns:
            index = selected[0] if selected else self.session.editor.column_count
            self.act(lambda e: e.insert_columns(index, count))
        else:
            index = selected[0] + self.session.offset if selected else len(self.session.editor.rows)
            self.act(lambda e: e.insert_rows(index, count))

    def delete(self, columns):
        if self.session is None:
            return
        selected = self.selected(columns)
        if not selected:
            messagebox.showinfo('選択', '列番号または行番号で選択してください。', parent=self.root)
            return
        if columns:
            self.act(lambda e: e.delete_columns(selected))
        else:
            indices = [r + self.session.offset for r in selected]
            self.act(lambda e: e.delete_rows(indices))

    def edit_header(self):
        if not self.session or not self.session.offset:
            messagebox.showinfo('見出し', '見出しありのファイルを開いてください。', parent=self.root)
            return
        column = simpledialog.askinteger('見出し編集', '列番号（1から）', minvalue=1, maxvalue=self.session.editor.column_count, parent=self.root)
        if column is None:
            return
        row = self.session.editor.rows[0]
        initial = row[column - 1] if column <= len(row) else ''
        value = simpledialog.askstring('見出し編集', '見出しの文字列', initialvalue=initial, parent=self.root)
        if value is not None:
            self.act(lambda e: e.set_cell(0, column - 1, value))

    def find(self):
        if self.session is None:
            return
        text = simpledialog.askstring('検索', '検索文字列（部分一致・大文字小文字を区別）', parent=self.root)
        if not text:
            return
        self.snapshot()
        found = self.session.editor.find(text)
        body = [(r - self.session.offset, c) for r, c in found if r >= self.session.offset]
        if body:
            self.sheet.select_cell(*body[0])
            self.sheet.see(*body[0])
        messagebox.showinfo('検索結果', f'{len(found)}セルに一致しました（見出しを含む）。\n本文の最初の一致を選択しました。' if body else f'{len(found)}セルに一致しました（見出しを含む）。', parent=self.root)

    def replace(self):
        if self.session is None:
            return
        old = simpledialog.askstring('一括置換', '検索文字列（見出しを含む全セルが対象）', parent=self.root)
        if not old:
            return
        new = simpledialog.askstring('一括置換', '置換後の文字列（空欄も可）', parent=self.root)
        if new is None:
            return
        self.act(lambda e: e.replace_all(old, new))


def main():
    root = tk.Tk()
    EditorGuiProbe(root)
    root.mainloop()


if __name__ == '__main__':
    main()
