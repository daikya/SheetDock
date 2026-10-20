"""CSV-only sort/filter integration. Widget data always contains every body row."""
import csv
import io
from tkinter import ttk
import tkinter as tk

from . import dialogs
from .gui import MultiFileApp, Session
from .diagnostics import record_failure


INPUT_CLASSES = {'Entry', 'TEntry', 'Text', 'TCombobox', 'Spinbox', 'TSpinbox'}


def history_shortcut(event, window_system):
    if event.state & 4 and (event.keysym.lower() == 'z' or (window_system == 'win32' and event.keycode == 90)):
        return bool(event.state & 1)
    return None


class CsvToolsApp(MultiFileApp):
    def csv_active(self):
        return self.active is not None and isinstance(self.active.session, Session)

    def csv_filtered(self):
        return self.csv_active() and self.active.session.editor.condition is not None

    def setup_csv_tools(self, root):
        self.tool_column = tk.StringVar(value='1')
        self.tool_mode = tk.StringVar(value='文字列')
        self.tool_descending = tk.BooleanVar(value=False)
        self.tool_operator = tk.StringVar(value='部分一致')
        self.tool_query = tk.StringVar()
        self._tools_context = None
        bar = self.csv_tools_bar = ttk.Frame(root)
        bar.grid(row=5, column=0, sticky='ew', padx=8, pady=3)
        ttk.Label(bar, text='CSV 対象列').pack(side='left')
        ttk.Spinbox(bar, from_=1, to=100000, textvariable=self.tool_column, width=5).pack(side='left', padx=3)
        ttk.Combobox(bar, textvariable=self.tool_mode, values=('文字列', '数値'), state='readonly', width=7).pack(side='left')
        ttk.Checkbutton(bar, text='降順', variable=self.tool_descending).pack(side='left')
        ttk.Button(bar, text='並べ替え', command=self.sort_csv).pack(side='left', padx=3)
        ttk.Combobox(bar, textvariable=self.tool_operator, values=('部分一致', '完全一致', '空', '空以外'), state='readonly', width=9).pack(side='left')
        ttk.Entry(bar, textvariable=self.tool_query, width=16).pack(side='left', padx=3)
        ttk.Button(bar, text='フィルター適用', command=self.filter_csv).pack(side='left')
        ttk.Button(bar, text='解除', command=self.clear_csv_filter).pack(side='left', padx=3)
        root.rowconfigure(5, weight=0)
        self.notebook.grid_configure(row=6)
        root.rowconfigure(6, weight=1)
        self.guidance_label.grid_configure(row=7)
        self.status_label.grid_configure(row=8)
        self._csv_history_tag = f'CsvHistory_{id(self)}'
        root.bind_class(self._csv_history_tag, '<KeyPress>', self.csv_key)
        self.install_csv_tag(root)
        self.refresh_csv_tools()

    def install_csv_tag(self, widget):
        tags = widget.bindtags()
        widget.bindtags((self._csv_history_tag,) + tuple(tag for tag in tags if tag != self._csv_history_tag))
        for child in widget.winfo_children():
            self.install_csv_tag(child)

    def model_history_key(self, event, redo):
        if event.widget.winfo_class() in INPUT_CLASSES:
            return None
        return self.history(redo)

    def csv_key(self, event):
        if not self.csv_active() or event.widget.winfo_class() in INPUT_CLASSES:
            return None
        redo = history_shortcut(event, self.root.tk.call('tk', 'windowingsystem'))
        if redo is not None:
            self.history(redo)
            return 'break'
        if self.csv_filtered() and event.state & 4 and event.keysym.lower() == 'v':
            self.paste_filtered()
            return 'break'
        return None

    def refresh_csv_tools(self):
        if not hasattr(self, 'csv_tools_bar'):
            return
        if not self.csv_active():
            self._tools_context = None
            self.csv_tools_bar.grid_remove()
            return
        self.csv_tools_bar.grid()
        context = (self.active, self.active.session)
        if context != self._tools_context:
            self._tools_context = context
            condition = self.session.editor.condition
            column, operator, query = condition or (0, 'contains', '')
            self.tool_column.set(str(column + 1))
            self.tool_operator.set({'contains': '部分一致', 'equals': '完全一致', 'empty': '空', 'nonempty': '空以外'}[operator])
            self.tool_query.set(query)
        filtered = self.csv_filtered()
        gated = {'行を挿入', '選択行を削除', '列を挿入', '選択列を削除', '見出し編集', '一括置換', '行の高さ'}
        for button in self.row_column_actions.winfo_children():
            if isinstance(button, ttk.Button) and button.cget('text') in gated:
                button.configure(state='disabled' if filtered else 'normal')
        self.header_toggle.configure(state='disabled' if filtered else 'normal')
        self.command_buttons['find'].configure(state='disabled' if filtered else 'normal')
        self.command_buttons['paste'].configure(state='normal' if filtered else 'disabled')
        self.command_buttons['undo'].configure(state='normal' if self.session.editor._undo else 'disabled')
        self.command_buttons['redo'].configure(state='normal' if self.session.editor._redo else 'disabled')
        self.guidance_label.configure(text=('フィルター中：保存は非表示行を含む全行。検索・構造変更等は解除後に操作してください。'
                                           if filtered else '並べ替えは行全体を移動しUndo対象。フィルターは表示だけを絞ります。'))

    def tool_action(self, action):
        if not self.csv_active():
            return
        try:
            self.snapshot()
            action()
            self.render()
            self.sheet.MT.focus_set()
        except (ValueError, TypeError, IndexError, OSError, csv.Error, tk.TclError) as error:
            record_failure('csv_sort_filter', error, warning=True)
            dialogs.showerror('CSV操作中止', str(error), parent=self.root)

    def sort_csv(self):
        column, mode, descending = self.tool_column.get(), self.tool_mode.get(), self.tool_descending.get()
        self.tool_action(lambda: self.session.editor.sort_rows(int(column) - 1,
                         mode={'文字列': 'text', '数値': 'number'}[mode], descending=descending))

    def filter_csv(self):
        column, operator, query = self.tool_column.get(), self.tool_operator.get(), self.tool_query.get()
        self.tool_action(lambda: self.session.editor.set_filter(int(column) - 1,
                         operator={'部分一致': 'contains', '完全一致': 'equals', '空': 'empty', '空以外': 'nonempty'}[operator], query=query))

    def clear_csv_filter(self):
        self.tool_action(lambda: self.session.editor.clear_filter())

    def paste_filtered(self):
        def action():
            selected = self.sheet.get_currently_selected()
            if self.sheet.display_rows() != self.session.visible_body_indices():
                self.apply_csv_filter(self.active)
                raise ValueError('入力で表示対象が変わりました。セルを選び直してください')
            if not selected:
                raise ValueError('表示中の本文セルを選択してください')
            values = list(csv.reader(io.StringIO(self.root.clipboard_get(), newline=''), delimiter='\t', strict=True))
            self.session.editor.paste_visible(selected.row, selected.column, values)
        self.tool_action(action)

    def create_page(self, tab):
        super().create_page(tab)
        if hasattr(self, '_csv_history_tag'):
            self.install_csv_tag(self.pages[tab][0])

    def cancel_filter_refresh(self, tab):
        sheet = self.pages[tab][1]
        job = getattr(sheet, '_csv_filter_job', None)
        if job:
            self.root.after_cancel(job)
            sheet._csv_filter_job = None

    def apply_csv_filter(self, tab):
        sheet, session = self.pages[tab][1], tab.session
        filtered = session.editor.condition is not None
        self.applying_widths = True
        try:
            sheet.display_rows(rows=session.visible_body_indices() if filtered else 'all',
                               all_rows_displayed=False if filtered else True, redraw=False)
            sheet.row_index([str(i + session.offset + 1) for i in range(len(session.editor.rows) - session.offset)] if filtered else [])
            if filtered:
                sheet.disable_bindings('paste', 'row_height_resize')
            else:
                sheet.enable_bindings('paste', 'row_height_resize')
            self.apply_row_heights(tab, sheet)
        finally:
            self.applying_widths = False
        if tab is self.active:
            self.format_status()
            self.refresh_csv_tools()

    def queue_csv_filter(self, tab):
        sheet = self.pages[tab][1]
        if getattr(sheet, '_csv_filter_job', None):
            return
        def flush():
            sheet._csv_filter_job = None
            if tab not in self.pages:
                return
            if sheet.MT.text_editor.open:
                sheet._csv_filter_job = self.root.after(80, flush)
                return
            if tab.session.editor.condition is not None:
                self.apply_csv_filter(tab)
        sheet._csv_filter_job = self.root.after_idle(flush)

    def render(self):
        if self.csv_active():
            self.cancel_filter_refresh(self.active)
            self.sheet.display_rows('all', redraw=False)
        super().render()
        if self.csv_active():
            self.apply_csv_filter(self.active)

    def modified(self, event=None):
        super().modified(event)
        if self.csv_filtered() and not self.rendering:
            self.queue_csv_filter(self.active)

    def snapshot(self):
        if self.csv_active() and self.sheet.MT.text_editor.open:
            # A notebook can hide the Text before its FocusOut handler runs.
            # Settle its captured display coordinate without depending on focus.
            row, column = self.sheet.MT.text_editor.coords
            data_row = self.sheet.displayed_row_to_data(row) + self.session.offset
            if data_row not in self.session.editor.visible_indices():
                raise ValueError('表示条件が変わりました。入力セルを選び直してください')
            value = self.sheet.get_text_editor_value()
            self.session.editor.set_cell(data_row, column, value)
            self.sheet.MT.hide_text_editor_and_dropdown()
            self.sheet.set_cell_data(data_row - self.session.offset, column, value=value, redraw=True)
        return super().snapshot()

    def apply_row_heights(self, tab, sheet):
        if isinstance(tab.session, Session) and tab.session.editor.condition is not None and not sheet.MT.all_rows_displayed:
            minimum, default = sheet.MT.min_row_height, sheet.MT.get_default_row_height()
            values = tab.session.editor.row_heights
            sheet.set_row_heights([max(minimum, values[i + tab.session.offset] or default)
                                   for i in sheet.display_rows()])
            sheet.refresh()
        else:
            if isinstance(tab.session, Session) and not sheet.MT.all_rows_displayed:
                # Inactive reload replaces the session before it reapplies heights.
                # Clear the previous session's native row map first.
                sheet.display_rows('all', redraw=False)
                sheet.row_index([])
            super().apply_row_heights(tab, sheet)

    def format_status(self, notice=''):
        super().format_status(notice)
        if self.csv_filtered():
            session = self.active.session
            self.status.set(self.status.get() + f' / フィルター: {len(session.visible_body_indices())}/{len(session.editor.rows) - session.offset}行 / 保存は全行')

    def require_unfiltered(self):
        if self.csv_filtered():
            dialogs.showinfo('フィルター中', 'この操作の前にCSVフィルターを解除してください。', parent=self.root)
            return False
        return True

    def insert(self, columns):
        if self.require_unfiltered():
            return super().insert(columns)

    def delete(self, columns):
        if self.require_unfiltered():
            return super().delete(columns)

    def edit_header(self):
        if self.require_unfiltered():
            return super().edit_header()

    def find(self):
        if self.require_unfiltered():
            return super().find()

    def search_action(self, operation):
        if self.require_unfiltered():
            return super().search_action(operation)

    def toggle_header(self):
        if self.require_unfiltered():
            return super().toggle_header()

    def set_row_height(self):
        if self.require_unfiltered():
            return super().set_row_height()

    def finish_row_resize(self, tab):
        if isinstance(tab.session, Session) and tab.session.editor.condition is not None:
            self.pages[tab][1]._sheetdock_height_drag = {}
            return
        result = super().finish_row_resize(tab)
        if tab is self.active:
            self.refresh_csv_tools()
        return result

    def page_resized(self, tab):
        result = super().page_resized(tab)
        if tab is self.active:
            self.refresh_csv_tools()
        return result

    def reload_tab(self, tab):
        previous = tab.session
        result = super().reload_tab(tab)
        if isinstance(tab.session, Session) and tab.session is not previous and tab in self.pages:
            self.cancel_filter_refresh(tab)
            self.apply_csv_filter(tab)
        return result
