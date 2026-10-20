"""File tabs shared by the existing CSV editor and the XLSX workbook pane."""
from pathlib import Path
from tkinter import filedialog, ttk

from . import dialogs as messagebox
from .diagnostics import record_failure
from .gui import MultiFileApp
from .storage import fingerprint
from .xlsx_gui import XlsxPane, XlsxSession, protection_reason


def is_xlsx(tab):
    return tab is not None and isinstance(tab.session, XlsxSession)


class SheetDockApp(MultiFileApp):
    def __init__(self, root):
        super().__init__(root)
        self.build_shared_layout(root)
        for keys, redo in [('<Control-z>', False), ('<Control-Z>', False),
                           ('<Control-Shift-z>', True), ('<Control-Shift-Z>', True)]:
            root.bind(keys, lambda event, redo=redo: self.history(redo))
        root.bind('<Control-S>', lambda event: (self.save_current(), 'break')[1])
        root.bind('<Control-f>', lambda event: (self.find(), 'break')[1])
        self.status.set('CSV / TSV / XLSXを開いてください。')

    def build_shared_layout(self, root):
        # Reserve toolbar/footer rows before assigning remaining space to the table.
        for widget in root.winfo_children():
            if widget.winfo_manager() == 'pack':
                widget.pack_forget()
        for widget in self.csv_options.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.pack_forget()
        for widget in self.csv_encoding_actions.winfo_children():
            if isinstance(widget, ttk.Checkbutton) and widget is not self.header_toggle:
                widget.pack_forget()
        self.file_actions = ttk.Frame(root)
        self.file_actions.grid(row=0, column=0, sticky='ew', padx=8, pady=4)
        ttk.Button(self.file_actions, text='開く', command=self.open).pack(side='left', padx=3)
        self.save_current_button = ttk.Button(self.file_actions, text='上書き保存', command=self.save_current, state='disabled')
        self.save_current_button.pack(side='left', padx=3)
        self.save_button = ttk.Button(self.file_actions, text='別名保存', command=self.save, state='disabled')
        self.save_button.pack(side='left', padx=3)
        self.close_button = ttk.Button(self.file_actions, text='タブを閉じる', command=self.close_tab, state='disabled')
        self.close_button.pack(side='left', padx=3)
        ttk.Checkbutton(self.file_actions, text='保存時にバックアップ', variable=self.backup_enabled,
                        command=self.change_backup).pack(side='left', padx=12)
        self.edit_actions = ttk.Frame(root)
        self.edit_actions.grid(row=1, column=0, sticky='ew', padx=8, pady=3)
        self.command_buttons = {}
        for key, label, command in [
            ('undo', 'Undo', lambda: self.history(False)),
            ('redo', 'Redo', lambda: self.history(True)),
            ('find', '検索・置換', self.find),
            ('edit', 'セル編集', lambda: self.xlsx_action('edit_cell')),
            ('paste', '貼り付け', lambda: self.xlsx_action('paste')),
            ('export', 'CSV出力', lambda: self.xlsx_action('export_csv')),
        ]:
            button = ttk.Button(self.edit_actions, text=label, command=command, state='disabled')
            button.pack(side='left', padx=3)
            self.command_buttons[key] = button
        self.csv_options.grid(row=2, column=0, sticky='ew', padx=8, pady=3)
        self.csv_encoding_actions.grid(row=3, column=0, sticky='ew', padx=8, pady=3)
        self.row_column_actions.grid(row=4, column=0, sticky='ew', padx=8, pady=3)
        self.notebook.grid(row=5, column=0, sticky='nsew', padx=8, pady=4)
        self.guidance_label.grid(row=6, column=0, sticky='w', padx=8)
        self.status_label.grid(row=7, column=0, sticky='ew', padx=8, pady=4)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(5, weight=1)
        root.minsize(min(1050, root.winfo_screenwidth()), min(650, max(400, root.winfo_screenheight()-60)))
        self.layout_ready = True
        self.update_commands()

    def xlsx_action(self, action):
        if is_xlsx(self.active):
            return getattr(self.pane(), action)()

    def update_commands(self):
        if not getattr(self, 'layout_ready', False):
            return
        csv_active = self.active is not None and not is_xlsx(self.active)
        xlsx_active = is_xlsx(self.active)
        pane = self.pane() if xlsx_active else None
        session = self.active.session if xlsx_active else None
        editable = xlsx_active and not session.probe.read_only
        selected = pane.selected_cell() if xlsx_active else None
        cell_editable = editable and selected is not None and not protection_reason(session.probe, session.sheet, selected)
        states = {'undo': csv_active or (editable and bool(session.undo_stack)),
                  'redo': csv_active or (editable and bool(session.redo_stack)),
                  'find': self.active is not None, 'edit': cell_editable,
                  'paste': cell_editable, 'export': xlsx_active}
        for name, button in self.command_buttons.items():
            button.configure(state='normal' if states[name] else 'disabled')
        can_save = csv_active or editable
        self.save_current_button.configure(state='normal' if can_save else 'disabled')
        self.save_button.configure(state='normal' if can_save else 'disabled')
        self.close_button.configure(state='normal' if self.active is not None else 'disabled')
        # Only CSV controls are contextual; global file actions stay in place.
        if xlsx_active:
            self.csv_options.grid_remove()
            self.csv_encoding_actions.grid_remove()
            self.row_column_actions.grid_remove()
            self.guidance_label.grid_remove()
        else:
            self.csv_options.grid()
            self.guidance_label.grid()
            if csv_active:
                self.csv_encoding_actions.grid()
                self.row_column_actions.grid()
            else:
                self.csv_encoding_actions.grid_remove()
                self.row_column_actions.grid_remove()

    def pane(self, tab=None):
        tab = self.active if tab is None else tab
        return self.pages[tab][1] if is_xlsx(tab) else None

    def finish_xlsx_edit(self, tab=None):
        pane = self.pane(tab)
        finish = getattr(pane, 'finish_inline_edit', None)
        return finish is None or finish()

    def open(self):
        if not self.finish_xlsx_edit():
            return
        selected = filedialog.askopenfilename(parent=self.root, filetypes=[
            ('CSV / TSV / XLSX', '*.csv *.tsv *.xlsx'), ('CSV / TSV', '*.csv *.tsv'),
            ('Excel workbook', '*.xlsx'), ('すべて', '*.*')])
        if selected:
            self.open_path(selected)

    def open_path(self, selected):
        if not self.finish_xlsx_edit():
            return
        suffix = Path(selected).suffix.lower()
        if suffix in {'.xls', '.xlsm', '.xlsb', '.xltx', '.xltm'}:
            messagebox.showerror('未対応形式', 'Excel形式は.xlsxだけに対応しています。', parent=self.root)
            return
        if suffix != '.xlsx':
            return super().open_path(selected)
        try:
            tab = self.documents.find(selected)
            if tab is None:
                if not self.confirm_large_file(selected):
                    return
                session = XlsxSession(selected)
                tab = self.documents.add(session.probe.path, session, session.stamp)
                try:
                    self.create_page(tab)
                except Exception:
                    self.documents.remove(tab)
                    raise
            self.notebook.select(self.pages[tab][0])
            self.activate(tab)
            self.remember_file(tab.source)
        except Exception as exc:
            record_failure('read', exc)
            messagebox.showerror('読込失敗', str(exc), parent=self.root)

    def create_page(self, tab):
        if not is_xlsx(tab):
            return super().create_page(tab)
        page = ttk.Frame(self.notebook)
        try:
            pane = XlsxPane(page, owner=self.root, backup_var=self.backup_enabled,
                            on_open=self.open, on_change=lambda: self.sync_xlsx(tab),
                            on_save=lambda: self.save_tab(tab),
                            on_save_as=lambda: self.save_as_tab(tab),
                            check_export=self.check_export_target,
                            on_backup_change=self.change_backup, on_selection=self.update_commands)
            pane.session = tab.session
            self.pages[tab] = (page, pane)
            self.notebook.add(page, text=tab.title)
        except Exception:
            self.pages.pop(tab, None)
            page.destroy()
            raise

    def sync_xlsx(self, tab):
        pane = self.pane(tab)
        tab.session = pane.session
        tab.source, tab.stamp = tab.session.probe.path, tab.session.stamp
        if tab is self.active:
            self.source = tab.source
            self.format_status()
        self.labels()

    def activate(self, tab):
        if tab is not self.active and not self.finish_xlsx_edit():
            self.switching = True
            try:
                self.notebook.select(self.pages[self.active][0])
            finally:
                self.switching = False
            return
        if not is_xlsx(tab):
            return super().activate(tab)
        if tab is self.active:
            return
        if self.active is not None:
            self.snapshot()
            self.persist_widths(self.active)
        if self.search_window is not None and self.search_window.winfo_exists():
            self.search_window.destroy()
        self.active, self.source = tab, tab.source
        # CSV commands cannot operate on an XLSX model or a stale CSV sheet.
        self.session = self.sheet = self.document = None
        pane = self.pane(tab)
        if not getattr(pane, '_sheetdock_ready', False):
            pane.render()
            pane._sheetdock_ready = True
        self.save_button.configure(state='disabled' if tab.session.probe.read_only else 'normal')
        self.format_status()
        self.labels()

    def labels(self):
        for tab, (page, _) in self.pages.items():
            text = tab.title
            if is_xlsx(tab) and tab.session.probe.read_only:
                text += ' [読取専用]'
            self.notebook.tab(page, text=text)
        csv_active = self.active is not None and not is_xlsx(self.active)
        self.header_toggle.configure(state='normal' if csv_active else 'disabled')
        self.header_display.set(self.session.has_header if csv_active else False)
        can_save = csv_active or (is_xlsx(self.active) and not self.active.session.probe.read_only)
        self.save_button.configure(state='normal' if can_save else 'disabled')
        if self.active is None:
            self.status.set('CSV / TSV / XLSXを開いてください。')
        self.update_commands()

    def format_status(self, notice=''):
        if is_xlsx(self.active):
            session = self.active.session
            mode = '読取専用' if session.probe.read_only else 'XLSX'
            dirty = 'あり' if session.dirty else 'なし'
            self.status.set(f'{notice} {session.probe.path} / {mode} / 未保存={dirty}')
            return
        return super().format_status(notice)

    def render(self):
        if is_xlsx(self.active):
            return self.pane().render()
        return super().render()

    def snapshot(self):
        if is_xlsx(self.active):
            return None
        return super().snapshot()

    def capture_widths(self, tab, persist=False):
        if not is_xlsx(tab):
            return super().capture_widths(tab, persist)

    def persist_widths(self, tab):
        if not is_xlsx(tab):
            return super().persist_widths(tab)

    def history(self, redo):
        if is_xlsx(self.active):
            return self.pane().history(redo)
        return super().history(redo)

    def find(self):
        if not self.finish_xlsx_edit():
            return
        if is_xlsx(self.active):
            self.pane().query_entry.focus_set()
            return
        return super().find()

    def replace(self):
        return self.find()

    def search_action(self, operation):
        if is_xlsx(self.active):
            self.find()
            return
        return super().search_action(operation)

    def reopen_encoding(self):
        if not is_xlsx(self.active):
            return super().reopen_encoding()

    def toggle_header(self):
        if not is_xlsx(self.active):
            return super().toggle_header()

    def save_as_tab(self, tab):
        if not self.finish_xlsx_edit(tab):
            return False
        if not is_xlsx(tab):
            return super().save_as_tab(tab)
        if tab.session.probe.read_only:
            return False

        selected = filedialog.asksaveasfilename(parent=self.root, initialdir=str(tab.source.parent),
            initialfile=tab.source.stem + '_edited.xlsx', defaultextension='.xlsx',
            filetypes=[('Excel workbook', '*.xlsx')])
        if not selected:
            return False
        try:
            self.documents.check_save_as(tab, selected)
            tab.session.save_as(selected)  # New destination only.
            self.pane(tab).render()
            self.remember_file(tab.source)
            return True
        except Exception as exc:
            record_failure('save', exc)
            messagebox.showerror('保存失敗', str(exc), parent=self.root)
            return False

    def validate_save_extension(self, tab, path):
        if not is_xlsx(tab) and Path(path).suffix.lower() in {'.xlsx', '.xls', '.xlsm', '.xlsb', '.xltx', '.xltm'}:
            raise ValueError('CSVの保存先にExcel形式の拡張子は指定できません。CSVからXLSXへの変換は未対応です。')

    def check_export_target(self, path):
        if self.documents.find(path) is not None:
            raise ValueError('CSV出力先は別のタブで開かれています。新しい保存先を選んでください。')

    def save_tab(self, tab):
        if not self.finish_xlsx_edit(tab):
            return False
        if not is_xlsx(tab):
            return super().save_tab(tab)
        if tab.session.probe.read_only:
            return False
        try:
            observed = fingerprint(tab.source)
            expected = tab.session.stamp
            if observed != expected:
                choice = messagebox.askyesnocancel('外部変更を検出',
                    f'{tab.source.name}は外部で変更または削除されました。\n'
                    'はい：編集中の内容で上書き\nいいえ：再読込\nキャンセル：保存しない', parent=self.root)
                if choice is None:
                    return False
                if choice is False:
                    return self.reload_tab(tab)
                expected = observed
            self.check_backup_target(tab.source)
            tab.session.save(backup=self.backup_is_enabled(), expected=expected)
            self.pane(tab).render()
            return True
        except Exception as exc:
            record_failure('save', exc)
            messagebox.showerror('保存失敗', str(exc), parent=self.root)
            return False

    def reload_tab(self, tab):
        if not self.finish_xlsx_edit(tab):
            return False
        if not is_xlsx(tab):
            return super().reload_tab(tab)
        if tab.session.dirty and not messagebox.askyesno('再読込の確認',
            f'{tab.source.name}の編集内容を破棄して再読込しますか？', parent=self.root):
            return False
        try:
            replacement = XlsxSession(tab.source)
            tab.session = replacement
            self.pane(tab).session = replacement
            self.pane(tab).render()
        except Exception as exc:
            record_failure('reload', exc)
            messagebox.showerror('再読込失敗', str(exc), parent=self.root)
        # Reload is not a successful save/close decision.
        return False

    def close_tab(self):
        if self.finish_xlsx_edit():
            return super().close_tab()

    def close(self):
        for tab in self.documents.tabs:
            if not self.finish_xlsx_edit(tab):
                return
        return super().close()
