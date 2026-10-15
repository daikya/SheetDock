"""SheetDockの表編集画面と複数ファイルタブ。"""

import copy
import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from . import dialogs as messagebox

from tksheet import Sheet
from .editor import Editor
from . import __version__
from . import dialogs as simpledialog
from .dialogs import show_placed
from .csv_io import CsvDocument, load_csv, save_csv
from .storage import fingerprint, load_checked, save_checked, ExternalChangeError
from .encoding import EncodingSelectionCancelled
from .search import SearchState, find_matches
from .settings import SettingsStore, initial_geometry, geometry_parts
from .tabs import path_key
from .diagnostics import configure_logging, record_failure, shutdown_logging
from .large_csv import large_file, large_table, table_size, FILE_WARNING_BYTES
from .icons import apply_icon


def split_rows(rows, use_header):
    """見出しを別表示にしても、元データを変更しない。"""
    copied = copy.deepcopy(rows)
    if use_header and copied:
        return copied[0], copied[1:]
    return None, copied


def build_document(template, header, displayed_rows):
    rows = copy.deepcopy(displayed_rows)
    if header is not None:
        rows.insert(0, copy.deepcopy(header))
    if any(not isinstance(cell, str) for row in rows for cell in row):
        raise TypeError('セルに文字列以外の値があります。保存を中止しました。')
    return CsvDocument(rows, template.format)


class CsvWindow:
    def __init__(self, root):
        self.root = root
        self.document = None
        self.header = None
        self.source = None
        self.last_saved_rows = None
        root.title('SheetDock')
        root.geometry('1100x700')
        root.protocol('WM_DELETE_WINDOW', self.close)
        self.encoding = tk.StringVar(value='自動')
        self.delimiter = tk.StringVar(value='カンマ')
        self.use_header = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value='読込条件を選んでCSVを開いてください。')
        controls = ttk.Frame(root)
        controls.pack(fill='x', padx=10, pady=10)
        ttk.Label(controls, text='次に開くCSV：文字コード').pack(side='left')
        ttk.Combobox(controls, textvariable=self.encoding, values=('自動', 'utf-8', 'cp932'), state='readonly', width=10).pack(side='left', padx=5)
        ttk.Label(controls, text='区切り').pack(side='left')
        ttk.Combobox(controls, textvariable=self.delimiter, values=('カンマ', 'タブ', 'セミコロン'), state='readonly', width=10).pack(side='left', padx=5)
        ttk.Checkbutton(controls, text='先頭行を見出しにする', variable=self.use_header).pack(side='left', padx=5)
        ttk.Button(controls, text='CSVを開く', command=self.open).pack(side='left', padx=5)
        self.save_button = ttk.Button(controls, text='別名保存', command=self.save, state='disabled')
        self.save_button.pack(side='left', padx=5)
        self.sheet = Sheet(root)
        self.sheet.pack(fill='both', expand=True, padx=10, pady=5)
        # 行列構造の変更は次の検証。見出しと本文の位置関係を固定する。
        self.sheet.enable_bindings(
            'single_select', 'drag_select', 'select_all', 'column_select', 'row_select',
            'column_width_resize', 'row_height_resize', 'arrowkeys',
            'copy', 'cut', 'paste', 'delete', 'undo', 'redo', 'edit_cell', 'find', 'replace',
        )
        self.sheet.set_options(paste_can_expand_x=False, paste_can_expand_y=False)
        actions = ttk.Frame(root)
        actions.pack(fill='x', padx=10, pady=5)
        ttk.Button(actions, text='Undo', command=self.sheet.undo).pack(side='left')
        ttk.Button(actions, text='Redo', command=self.sheet.redo).pack(side='left', padx=5)
        ttk.Button(actions, text='検索・置換', command=lambda: self.sheet.open_find(focus=True)).pack(side='left')
        ttk.Label(root, text='見出しは表示専用。読込条件の変更は次に開くファイルから適用します。').pack(anchor='w', padx=10)
        ttk.Label(root, textvariable=self.status, wraplength=1050).pack(anchor='w', padx=10, pady=10)

    def snapshot(self):
        # 保存ボタン操作時に編集中のセル値を確定する。
        self.sheet.close_text_editor(set_data=True)
        return build_document(self.document, self.header, self.sheet.get_sheet_data())

    def may_discard(self):
        if self.document is None:
            return True
        try:
            changed = self.snapshot().rows != self.last_saved_rows
        except Exception as exc:
            record_failure('confirm', exc)
            messagebox.showerror('確認失敗', str(exc), parent=self.root)
            return False
        return not changed or messagebox.askyesno('未保存の編集', '未保存の編集を破棄して続けますか？', parent=self.root)

    def open(self):
        selected = filedialog.askopenfilename(parent=self.root, filetypes=[('CSV / TSV', '*.csv *.tsv'), ('すべて', '*.*')])
        if not selected:
            return
        try:
            delimiter = {'カンマ': ',', 'タブ': '\t', 'セミコロン': ';'}[self.delimiter.get()]
            loaded = load_csv(selected, encoding=self.encoding.get(), delimiter=delimiter)
        except Exception as exc:
            record_failure('read', exc)
            messagebox.showerror('読込失敗', str(exc), parent=self.root)
            return
        if not self.may_discard():
            return
        header, rows = split_rows(loaded.rows, self.use_header.get())
        self.sheet.set_sheet_data(rows, delete_options=True)
        self.sheet.headers(header if header is not None else [])
        self.sheet.reset_undos()
        self.document, self.header, self.source = loaded, header, Path(selected)
        self.last_saved_rows = copy.deepcopy(loaded.rows)
        self.save_button.configure(state='normal')
        fmt = loaded.format
        newline = {'\r\n': 'CRLF', '\n': 'LF', '\r': 'CR'}[fmt.newline]
        self.status.set(f'{selected} / {fmt.encoding} / BOM={fmt.bom} / {newline} / {len(loaded.rows)}行（見出しを含む）')

    def save(self):
        if self.document is None:
            return
        try:
            edited = self.snapshot()
        except Exception as exc:
            record_failure('save', exc)
            messagebox.showerror('保存前確認失敗', str(exc), parent=self.root)
            return
        selected = filedialog.asksaveasfilename(parent=self.root, initialdir=str(self.source.parent), initialfile=self.source.stem + '_edited' + self.source.suffix, defaultextension=self.source.suffix or '.csv', filetypes=[('CSV / TSV', '*.csv *.tsv'), ('すべて', '*.*')])
        if not selected:
            return
        target = Path(selected)
        same_path = os.path.normcase(str(target.resolve())) == os.path.normcase(str(self.source.resolve()))
        if target.exists() and self.source.exists():
            same_path = same_path or os.path.samefile(target, self.source)
        if same_path:
            messagebox.showerror('別名保存', '別名保存では読込元と異なる保存先を選んでください。', parent=self.root)
            return
        try:
            save_csv(edited, target, backup=True)
            # 別名出力なので読込元は変えず、保存済みの編集内容だけ更新する。
            self.last_saved_rows = copy.deepcopy(edited.rows)
        except Exception as exc:
            record_failure('save', exc)
            messagebox.showerror('保存失敗', str(exc), parent=self.root)
            return
        self.status.set(f'保存・再読込検証成功：{target}')
        messagebox.showinfo('保存完了', '別名保存と再読込での一致確認が完了しました。', parent=self.root)

    def close(self):
        if self.may_discard():
            self.root.destroy()



class Session:
    def __init__(self, document, use_header):
        self.format = document.format
        self.has_header = use_header
        self.editor = Editor(document.rows)
        self.search = SearchState()

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

    def set_header_view(self, enabled):
        if type(enabled) is not bool:
            raise TypeError('見出し表示は有効・無効で指定してください。')
        self.has_header = enabled


class SheetDockApp(CsvWindow):
    def __init__(self, root):
        self.session = None
        self.rendering = False
        super().__init__(root)
        root.title(f'SheetDock {__version__}')
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
            self.status.set(f'{self.source} / 未保存={self.session.editor.dirty}')
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
            record_failure('edit', exc)
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
            record_failure('operation', exc)
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


class MultiFileApp(SheetDockApp):
    def __init__(self, root):
        from .tabs import Documents
        self.documents = Documents()
        self.active = None
        self.pages = {}
        self.switching = False
        self.search_window = None
        self.applying_widths = False
        super().__init__(root)
        old_sheet = self.sheet
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5, before=old_sheet)
        old_sheet.destroy()
        self.sheet = None
        self.notebook.bind('<<NotebookTabChanged>>', self.tab_changed)
        controls = root.winfo_children()[0]
        ttk.Button(controls, text='保存', command=self.save_current).pack(side='left', padx=5)
        encoding_actions = ttk.Frame(root)
        encoding_actions.pack(fill='x', padx=10, before=self.notebook)
        ttk.Button(encoding_actions, text='文字コードで再読込', command=self.reopen_encoding).pack(side='left', padx=5)
        self.header_display = tk.BooleanVar(value=False)
        self.header_toggle = ttk.Checkbutton(encoding_actions, text='現在のタブの先頭行を見出しにする', variable=self.header_display, command=self.toggle_header, state='disabled')
        self.header_toggle.pack(side='left', padx=12)
        root.bind('<Control-s>', lambda event: (self.save_current(), 'break')[1])
        ttk.Button(controls, text='タブを閉じる', command=self.close_tab).pack(side='left', padx=5)
        self.settings_store = SettingsStore()
        self.settings = self.settings_store.load()
        warning = self.settings_store.warning
        self.backup_enabled = tk.BooleanVar(value=self.settings.backup_enabled)
        ttk.Checkbutton(encoding_actions, text='保存時にバックアップ', variable=self.backup_enabled, command=self.change_backup).pack(side='left', padx=12)
        menubar = tk.Menu(root)
        file_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label='ファイル', menu=file_menu)
        file_menu.add_command(label='開く', command=self.open)
        self.recent_menu = tk.Menu(file_menu, tearoff=False)
        file_menu.add_cascade(label='最近使ったファイル', menu=self.recent_menu)
        root.configure(menu=menubar)
        self.refresh_recent_menu()
        self.normal_geometry = None
        self.restoring_geometry = True
        root.bind('<Configure>', self.remember_geometry, add='+')
        root.after_idle(self.restore_geometry)
        if warning:
            record_failure('settings_read', warning=True)
            messagebox.showwarning('設定の読込', warning, parent=root)

    def load_widths(self, tab):
        store = getattr(self, 'settings_store', None)
        if store is not None:
            tab.session.editor.set_column_widths(store.load().file_widths.get(path_key(tab.source), []))

    def capture_widths(self, tab, persist=False):
        if self.applying_widths or tab not in self.pages:
            return
        tab.session.editor.set_column_widths(self.pages[tab][1].get_column_widths())
        if persist:
            self.persist_widths(tab)

    def persist_widths(self, tab):
        store = getattr(self, 'settings_store', None)
        if store is None or tab.session.editor.dirty:
            return
        try:
            self.settings = store.update(widths_file=tab.source, column_widths=tab.session.editor.column_widths)
        except Exception as exc:
            record_failure('widths_save', exc, warning=True)
            messagebox.showwarning('列幅保存失敗', str(exc), parent=self.root)

    def page_resized(self, tab):
        if not self.rendering:
            self.capture_widths(tab, persist=True)

    def backup_is_enabled(self):
        variable = getattr(self, 'backup_enabled', None)
        return variable.get() if variable is not None else True

    def confirm_large_file(self, path):
        if not large_file(path):
            return True
        record_failure('large_file', warning=True)
        size = Path(path).stat().st_size / 1024**2
        return messagebox.askyesno('大容量CSVの読込', f'ファイルサイズは約{size:.1f}MiBです。\n読込・表示・編集に時間とメモリを使う場合があります。\n読み込みますか？\nいいえを選ぶと現在の編集を維持して中止します。', parent=self.root)

    def confirm_large_table(self, path, document):
        # サイズによる警告を承認済みなら、セル数による警告は重ねない。
        if large_file(path) or not large_table(document.rows):
            return True
        record_failure('large_table', warning=True)
        rows, cells = table_size(document.rows)
        return messagebox.askyesno('大きな表の表示', f'{rows:,}行・{cells:,}セルを含むCSVです。\n表示・編集に時間とメモリを使う場合があります。\n表として開きますか？\nいいえを選ぶと現在の編集を維持して中止します。', parent=self.root)

    def change_backup(self):
        try:
            self.settings = self.settings_store.update(backup_enabled=self.backup_enabled.get())
        except Exception as exc:
            record_failure('settings_save', exc)
            # 永続化できなかった設定はUIでも戻す。
            self.backup_enabled.set(self.settings.backup_enabled)
            messagebox.showerror('設定保存失敗', str(exc), parent=self.root)

    def remember_file(self, path):
        if not hasattr(self, 'settings_store'):
            return
        try:
            self.settings = self.settings_store.update(recent_file=path)
            self.refresh_recent_menu()
        except Exception as exc:
            record_failure('recent_save', exc, warning=True)
            messagebox.showwarning('履歴保存失敗', str(exc), parent=self.root)

    def refresh_recent_menu(self):
        self.recent_menu.delete(0, 'end')
        if not self.settings.recent_files:
            self.recent_menu.add_command(label='履歴はありません', state='disabled')
        for i, path in enumerate(self.settings.recent_files, 1):
            self.recent_menu.add_command(label=f'{i} {Path(path).name} — {Path(path).parent}', command=lambda p=path: self.open_path(p))

    def remember_geometry(self, event):
        if not self.restoring_geometry and event.widget is self.root and self.root.state() == 'normal':
            value = self.root.geometry()
            if geometry_parts(value) is not None:
                self.normal_geometry = value

    def restore_geometry(self):
        # メニュー・表などの初期レイアウトを終えてから配置し、初回の自動配置を避ける。
        self.root.update_idletasks()
        geometry = initial_geometry(self.settings.window_geometry, self.root.winfo_screenwidth(), self.root.winfo_screenheight())
        self.root.geometry(geometry)
        self.root.deiconify()
        self.root.update_idletasks()
        self.root.geometry(geometry)
        self.normal_geometry = geometry
        self.restoring_geometry = False

    def persist_geometry(self):
        try:
            self.root.update_idletasks()
            geometry = self.root.geometry() if self.root.state() == 'normal' else self.normal_geometry
            if geometry_parts(geometry) is None:
                geometry = self.normal_geometry
            if geometry_parts(geometry) is not None:
                self.settings = self.settings_store.update(window_geometry=geometry)
        except Exception as exc:
            record_failure('geometry_save', exc, warning=True)
            messagebox.showwarning('画面設定の保存失敗', str(exc), parent=self.root)

    def requested_encoding(self):
        return 'auto' if self.encoding.get() == '自動' else self.encoding.get()

    def find(self):
        self.open_search()

    def replace(self):
        self.open_search()

    def open_search(self):
        if self.active is None:
            return
        if self.search_window is not None and self.search_window.winfo_exists():
            self.search_window.lift()
            return
        window = tk.Toplevel(self.root)
        window.withdraw()
        self.search_window = window
        window.title('検索・置換')
        window.transient(self.root)
        self.search_text = tk.StringVar(value='')
        self.replace_text = tk.StringVar(value='')
        self.search_result = tk.StringVar(value='現在のタブ全体を検索します（見出しを含む）。')
        fields = ttk.Frame(window, padding=12)
        fields.pack(fill='x')
        ttk.Label(fields, text='検索文字列').grid(row=0, column=0, sticky='w', pady=4)
        entry = ttk.Entry(fields, textvariable=self.search_text, width=40)
        entry.grid(row=0, column=1, padx=8, pady=4)
        ttk.Label(fields, text='置換後').grid(row=1, column=0, sticky='w', pady=4)
        ttk.Entry(fields, textvariable=self.replace_text, width=40).grid(row=1, column=1, padx=8, pady=4)
        actions = ttk.Frame(window, padding=8)
        actions.pack(fill='x')
        for label, operation in [('前へ', 'previous'), ('次へ', 'next'), ('1件置換', 'one'), ('すべて置換', 'all')]:
            ttk.Button(actions, text=label, command=lambda op=operation: self.search_action(op)).pack(side='left', padx=4)
        ttk.Label(window, textvariable=self.search_result, wraplength=480).pack(fill='x', padx=12, pady=10)
        ttk.Label(window, text='部分一致・大文字小文字を区別。最後まで進むと先頭に戻ります。').pack(padx=12, pady=5)
        window.bind('<Return>', lambda event: self.search_action('next'))
        window.bind('<Escape>', lambda event: window.destroy())
        show_placed(window, self.root, 'search')
        entry.focus_set()

    def search_action(self, operation):
        if self.active is None:
            self.search_result.set('CSVを開いてください。')
            return
        try:
            self.snapshot()
            text = self.search_text.get()
            editor = self.session.editor
            state = self.session.search
            notice = ''
            if operation == 'all':
                count = editor.replace_all(text, self.replace_text.get())
                state.reset()
                self.render()
                self.search_result.set(f'{self.source.name}：{count}箇所を置換しました。Undoで戻せます。')
                return
            if operation == 'one':
                changed = state.replace_current(editor, text, self.replace_text.get())
                self.render()
                notice = '1箇所を置換しました。' if changed else '変更はありません。'
            match = state.navigate(editor, text, backwards=operation == 'previous')
            if match is None:
                self.search_result.set(f'{self.source.name}：{notice}一致はありません。')
                return
            matches = find_matches(editor.rows, text)
            number = matches.index(match) + 1
            offset = self.session.offset
            if match.row < offset:
                self.sheet.select_column(match.column)
                if self.session.view()[1]:
                    self.sheet.see(0, match.column)
                location = f'見出し・{match.column + 1}列'
            else:
                self.sheet.select_cell(match.row - offset, match.column)
                self.sheet.see(match.row - offset, match.column)
                location = f'本文{match.row - offset + 1}行・{match.column + 1}列'
            self.search_result.set(f'{self.source.name}：{notice}{number}/{len(matches)}箇所 — {location}・セル内{match.start + 1}文字目')
        except Exception as exc:
            record_failure('search', exc, warning=True)
            self.search_result.set(str(exc))

    def toggle_header(self):
        if self.active is None:
            return
        enabled = self.header_display.get()
        try:
            # 旧表示のセル編集中データを確定してから表示モードを変える。
            self.snapshot()
            self.session.set_header_view(enabled)
            self.render()
        except Exception as exc:
            record_failure('header_toggle', exc)
            self.header_display.set(self.session.has_header)
            messagebox.showerror('見出し切替失敗', str(exc), parent=self.root)

    def choose_encoding(self, decision):
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title('文字コードを選択')
        dialog.transient(self.root)
        ttk.Label(dialog, text='UTF-8とCP932の両方で読めます。日本語が正しく見える候補を選んでください。').pack(padx=12, pady=10)
        selected = tk.StringVar(value='utf-8')
        previews = ttk.Frame(dialog)
        previews.pack(fill='both', expand=True, padx=12)
        for encoding in decision.candidates:
            column = ttk.Frame(previews)
            column.pack(side='left', fill='both', expand=True, padx=5)
            ttk.Radiobutton(column, text=encoding, value=encoding, variable=selected).pack(anchor='w')
            text = tk.Text(column, width=38, height=12, wrap='none')
            text.insert('1.0', decision.previews[encoding])
            text.configure(state='disabled')
            text.pack(fill='both', expand=True)
        result = [None]
        def accept():
            result[0] = selected.get()
            dialog.destroy()
        buttons = ttk.Frame(dialog)
        buttons.pack(pady=10)
        ttk.Button(buttons, text='選択して開く', command=accept).pack(side='left', padx=5)
        ttk.Button(buttons, text='キャンセル', command=dialog.destroy).pack(side='left', padx=5)
        show_placed(dialog, self.root, 'encoding')
        dialog.grab_set()
        self.root.wait_window(dialog)
        return result[0]

    def format_status(self, notice=''):
        if not self.active:
            return
        fmt = self.session.format
        newline = {'\r\n': 'CRLF', '\n': 'LF', '\r': 'CR'}[fmt.newline]
        dirty = 'あり' if self.session.editor.dirty else 'なし'
        prefix = f'{notice}：' if notice else ''
        self.status.set(f'{prefix}{self.source} / {fmt.encoding} / BOM={"あり" if fmt.bom else "なし"} / {newline} / 未保存={dirty}')

    def reopen_encoding(self):
        tab = self.active
        if tab is None:
            return
        self.snapshot()
        if tab.session.editor.dirty and not messagebox.askyesno('再読込の確認', '未保存の編集を破棄して、選択した文字コードで開き直しますか？', parent=self.root):
            return
        try:
            if not self.confirm_large_file(tab.source):
                return False
            fmt = tab.session.format
            document, stamp = load_checked(tab.source, encoding=self.requested_encoding(), delimiter=fmt.delimiter, choose_encoding=self.choose_encoding)
            if not self.confirm_large_table(tab.source, document):
                return False
            replacement = Session(document, tab.session.has_header)
        except EncodingSelectionCancelled:
            return
        except Exception as exc:
            record_failure('reload', exc)
            messagebox.showerror('再読込失敗', str(exc), parent=self.root)
            return
        tab.session, tab.stamp = replacement, stamp
        self.load_widths(tab)
        self.session, self.document = replacement, document
        self.render()

    def create_page(self, tab):
        page = ttk.Frame(self.notebook)
        sheet = Sheet(page)
        sheet.pack(fill='both', expand=True)
        sheet.enable_bindings(
            'single_select', 'drag_select', 'select_all', 'column_select', 'row_select',
            'column_width_resize', 'row_height_resize', 'arrowkeys',
            'copy', 'cut', 'paste', 'delete', 'edit_cell',
        )
        sheet.set_options(paste_can_expand_x=False, paste_can_expand_y=False)
        sheet.bind('<<SheetModified>>', lambda event: self.page_modified(tab))
        sheet.extra_bindings('column_width_resize', lambda event: self.page_resized(tab))
        self.pages[tab] = (page, sheet)
        self.notebook.add(page, text=tab.title)

    def page_modified(self, tab):
        if tab is self.active:
            self.modified()

    def labels(self):
        for tab, (page, _) in self.pages.items():
            self.notebook.tab(page, text=tab.title)
        self.header_toggle.configure(state='normal' if self.active is not None else 'disabled')
        self.header_display.set(self.session.has_header if self.active is not None else False)

    def render(self):
        if self.active is None:
            return
        super().render()
        self.applying_widths = True
        try:
            self.sheet.set_column_widths(self.session.editor.column_widths)
        finally:
            self.applying_widths = False
        self.format_status()
        self.labels()

    def modified(self, event=None):
        if self.sheet is None:
            return
        super().modified(event)
        self.format_status()
        self.labels()

    def activate(self, tab):
        if tab is self.active:
            return
        if self.active is not None:
            self.snapshot()
            self.persist_widths(self.active)
        self.active = tab
        self.session = tab.session
        self.source = tab.source
        self.document = self.session.snapshot()
        self.last_saved_rows = None
        self.sheet = self.pages[tab][1]
        self.save_button.configure(state='normal')
        # 初回のみデータを設定。タブ切替で選択・列幅をリセットしない。
        if not getattr(self.sheet, '_sheetdock_ready', False):
            self.render()
            self.sheet._sheetdock_ready = True
        else:
            self.header = self.session.view()[0]
            self.format_status()
        self.labels()

    def tab_changed(self, event=None):
        if self.switching:
            return
        selected = self.notebook.select()
        for tab, (page, _) in self.pages.items():
            if str(page) == selected:
                self.activate(tab)
                return

    def open(self):
        selected = filedialog.askopenfilename(parent=self.root, filetypes=[('CSV / TSV', '*.csv *.tsv'), ('すべて', '*.*')])
        if selected:
            self.open_path(selected)

    def open_path(self, selected):
        try:
            tab = self.documents.find(selected)
            if tab is None:
                if not self.confirm_large_file(selected):
                    return
                delimiter = {'カンマ': ',', 'タブ': '\t', 'セミコロン': ';'}[self.delimiter.get()]
                loaded, stamp = load_checked(selected, encoding=self.requested_encoding(), delimiter=delimiter, choose_encoding=self.choose_encoding)
                if not self.confirm_large_table(selected, loaded):
                    return
                tab = self.documents.add(selected, Session(loaded, self.use_header.get()), stamp)
                self.load_widths(tab)
                self.create_page(tab)
            self.notebook.select(self.pages[tab][0])
            self.activate(tab)
            self.remember_file(tab.source)
        except EncodingSelectionCancelled:
            return
        except Exception as exc:
            record_failure('read', exc)
            messagebox.showerror('読込失敗', str(exc), parent=self.root)

    def snapshot(self):
        if self.active is None:
            return None
        self.capture_widths(self.active)
        return super().snapshot()

    def save_as_tab(self, tab):
        if tab is self.active:
            self.snapshot()
        selected = filedialog.asksaveasfilename(parent=self.root, initialdir=str(tab.source.parent), initialfile=tab.source.stem + '_edited' + tab.source.suffix, defaultextension=tab.source.suffix or '.csv', filetypes=[('CSV / TSV', '*.csv *.tsv'), ('すべて', '*.*')])
        if not selected:
            return False
        try:
            self.documents.check_save_as(tab, selected)
            self.check_backup_target(selected)
            expected = fingerprint(selected)
            stamp = save_checked(tab.session.snapshot(), selected, expected, backup=self.backup_is_enabled())
            self.documents.saved_as(tab, selected)
            tab.stamp = stamp
            self.remember_file(tab.source)
            self.persist_widths(tab)
        except Exception as exc:
            record_failure('save', exc)
            messagebox.showerror('保存失敗', str(exc), parent=self.root)
            return False
        if tab is self.active:
            self.source = tab.source
            self.last_saved_rows = tab.session.editor.rows
            self.format_status('保存完了')
        self.labels()
        return True

    def save(self):
        if self.active is not None:
            return self.save_as_tab(self.active)
        return False

    def save_current(self):
        if self.active is not None:
            return self.save_tab(self.active)
        return False

    def check_backup_target(self, target):
        if not self.backup_is_enabled():
            return
        path = Path(target)
        if path.exists() and self.documents.find(path.with_name(path.name + '.bak')) is not None:
            raise ValueError('バックアップ先のファイルが別のタブで開かれています。そのタブを閉じてから保存してください。')

    def reload_tab(self, tab):
        if tab.session.editor.dirty and not messagebox.askyesno('再読込の確認', f'{tab.source.name}の未保存の編集を破棄して再読込しますか？', parent=self.root):
            return False
        try:
            if not self.confirm_large_file(tab.source):
                return False
            fmt = tab.session.format
            document, stamp = load_checked(tab.source, encoding=fmt.encoding, delimiter=fmt.delimiter)
            if not self.confirm_large_table(tab.source, document):
                return False
            replacement = Session(document, tab.session.has_header)
        except Exception as exc:
            record_failure('reload', exc)
            messagebox.showerror('再読込失敗', str(exc), parent=self.root)
            return False
        tab.session, tab.stamp = replacement, stamp
        self.load_widths(tab)
        page, sheet = self.pages[tab]
        if tab is self.active:
            self.session = replacement
            self.document = document
            self.render()
        else:
            header, body = replacement.view()
            sheet.set_sheet_data(body, delete_options=True)
            sheet.headers(header if header is not None else [])
            sheet.reset_undos()
            sheet.set_column_widths(tab.session.editor.column_widths)
        self.labels()
        # 再読込は保存の代わりではないので、閉じる操作はここで止める。
        return False

    def save_tab(self, tab):
        if tab is self.active:
            self.snapshot()
        try:
            current = fingerprint(tab.source)
            if current != tab.stamp:
                record_failure('external_change', warning=True)
                choice = messagebox.askyesnocancel('外部変更を検出', f'{tab.source.name}は読込後に外部で変更または削除されました。\nはい：編集中の内容で上書き\nいいえ：外部の内容を再読込\nキャンセル：保存しない', parent=self.root)
                if choice is None:
                    return False
                if choice is False:
                    return self.reload_tab(tab)
                expected = current
            else:
                expected = tab.stamp
            self.check_backup_target(tab.source)
            tab.stamp = save_checked(tab.session.snapshot(), tab.source, expected, backup=self.backup_is_enabled())
            tab.session.editor.mark_saved()
            self.persist_widths(tab)
        except Exception as exc:
            record_failure('save', exc)
            messagebox.showerror('保存失敗', str(exc), parent=self.root)
            return False
        self.labels()
        if tab is self.active:
            self.format_status('保存完了')
        return True

    def decide_close(self, tab):
        result = messagebox.askyesnocancel('未保存の編集', f'{tab.source.name}の編集を保存して閉じますか？\nはい：保存／いいえ：破棄／キャンセル：閉じない', parent=self.root)
        return 'save' if result is True else 'discard' if result is False else 'cancel'

    def close_tab(self):
        tab = self.active
        if tab is None:
            return
        self.snapshot()
        if not self.documents.may_close(tab, self.decide_close, self.save_tab):
            return
        self.persist_widths(tab)
        page, sheet = self.pages.pop(tab)
        self.switching = True
        try:
            self.notebook.forget(page)
            page.destroy()
            self.documents.remove(tab)
            self.active = self.session = self.document = self.sheet = self.source = None
        finally:
            self.switching = False
        if self.documents.tabs:
            selected = self.notebook.select()
            next_tab = next((t for t, (p, _) in self.pages.items() if str(p) == selected), self.documents.tabs[0])
            self.notebook.select(self.pages[next_tab][0])
            self.activate(next_tab)
        else:
            self.save_button.configure(state='disabled')
            self.status.set('読込条件を選んでCSVを開いてください。')
            self.labels()

    def close(self):
        if self.active is not None:
            self.snapshot()
        # 全タブの確認が完了するまでタブを削除しない。
        for tab in self.documents.tabs:
            if not self.documents.may_close(tab, self.decide_close, self.save_tab):
                return
        for tab in self.documents.tabs:
            self.capture_widths(tab)
            self.persist_widths(tab)
        self.persist_geometry()
        self.root.destroy()


def main():
    log_path = configure_logging()
    try:
        root = tk.Tk()
        root.withdraw()
        apply_icon(root)
        app = MultiFileApp(root)
        def callback_error(exc_type, value, traceback):
            record_failure('gui_callback', value)
            messagebox.showerror('画面処理のエラー', '画面処理でエラーが発生しました。編集中の内容を確認してください。', parent=root)
        root.report_callback_exception = callback_error
        if log_path is None:
            root.after_idle(lambda: messagebox.showwarning('ログの保存', 'ログを保存できません。アプリはログなしで続行します。', parent=root))
        root.mainloop()
    except Exception as exc:
        record_failure('startup', exc)
        raise
    finally:
        shutdown_logging()


if __name__ == '__main__':
    main()
