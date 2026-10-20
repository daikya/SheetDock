"""XLSX sort/filter controls over a full-data native sheet view."""
import tkinter as tk
from tkinter import ttk

from . import dialogs
from .csv_tools_gui import history_shortcut, INPUT_CLASSES
from .diagnostics import record_failure
from .xlsx_axis import VIEW_ROWS, VIEW_COLUMNS
from .xlsx_extent import extent
from .xlsx_inline import settled
from .xlsx_search import XlsxSearch


class XlsxToolsPane:
    def setup_xlsx_tools(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill='x', pady=(4, 0))
        self.sort_column = tk.StringVar(value='1')
        self.sort_descending = tk.BooleanVar(value=False)
        self.sort_header = tk.BooleanVar(value=True)
        self.filter_operator = tk.StringVar(value='部分一致')
        self.filter_query = tk.StringVar()
        self._xlsx_tools_context = None
        ttk.Label(bar, text='XLSX 対象列').pack(side='left')
        ttk.Spinbox(bar, from_=1, to=VIEW_COLUMNS, textvariable=self.sort_column, width=4).pack(side='left', padx=3)
        self.sort_header_button = ttk.Checkbutton(bar, text='先頭行は見出し', variable=self.sort_header,
                                                 command=self.change_sort_header)
        self.sort_header_button.pack(side='left')
        ttk.Checkbutton(bar, text='降順', variable=self.sort_descending).pack(side='left')
        self.sort_button = ttk.Button(bar, text='型順に並べ替え', command=self.sort_xlsx)
        self.sort_button.pack(side='left', padx=3)
        ttk.Combobox(bar, textvariable=self.filter_operator, state='readonly', width=8,
                     values=('部分一致', '完全一致', '空', '空以外')).pack(side='left')
        ttk.Entry(bar, textvariable=self.filter_query, width=12).pack(side='left', padx=3)
        self.filter_button = ttk.Button(bar, text='フィルター適用', command=self.filter_xlsx)
        self.filter_button.pack(side='left')
        ttk.Button(bar, text='解除', command=self.clear_xlsx_filter).pack(side='left', padx=3)
        self._xlsx_key_tag = f'XlsxHistory_{id(self)}'
        self.root.bind_class(self._xlsx_key_tag, '<KeyPress>', self.xlsx_key)
        def install(widget):
            widget.bindtags((self._xlsx_key_tag,) + widget.bindtags())
            for child in widget.winfo_children():
                install(child)
        # Only this workbook pane; global CSV and file buttons keep their routing.
        install(self.sheet_widget.master)

    def xlsx_key(self, event):
        if event.widget.winfo_class() in INPUT_CLASSES:
            return None
        system = self.root.tk.call('tk', 'windowingsystem')
        redo = history_shortcut(event, system)
        if redo is not None:
            self.history(redo)
            return 'break'
        if event.state & 4 and (event.keysym.lower() == 'v' or (system == 'win32' and event.keycode == 86)):
            self.paste()
            return 'break'
        return None

    def sync_xlsx_tools(self):
        if not hasattr(self, 'sort_column'):
            return
        session = self.session
        context = (session, session.sheet_name)
        if self._xlsx_tools_context != context:
            self._xlsx_tools_context = context
            column, operator, query, header = session.filters.get(session.sheet_name,
                (1, 'contains', '', session.filter_headers.get(session.sheet_name, True)))
            self.sort_column.set(str(column))
            self.filter_operator.set({'contains': '部分一致', 'equals': '完全一致', 'empty': '空', 'nonempty': '空以外'}[operator])
            self.filter_query.set(query)
            self.sort_header.set(header)
        rows, columns = extent(session.sheet)
        bounded = rows <= VIEW_ROWS and columns <= VIEW_COLUMNS
        self.sort_button.configure(state='disabled' if session.probe.read_only or session.filtered or not bounded else 'normal')
        self.filter_button.configure(state='normal' if bounded else 'disabled')
        self.sort_header_button.configure(state='disabled' if session.filtered else 'normal')

    def change_sort_header(self):
        if self.session:
            if self.session.filtered:
                self.sort_header.set(self.session.filters[self.session.sheet_name][3])
            else:
                self.session.filter_headers[self.session.sheet_name] = self.sort_header.get()

    def apply_xlsx_filter(self):
        session, widget = self.session, self.sheet_widget
        indices = session.displayed_indices()
        widget.display_rows(indices if session.filtered else 'all',
                            all_rows_displayed=not session.filtered, redraw=False)
        widget.row_index([str(row + 1) for row in range(min(extent(session.sheet)[0], VIEW_ROWS))]
                         if session.filtered else [])
        if session.filtered:
            widget.disable_bindings('row_height_resize')
            heights = session.row_heights(self._point_scale, self._row_minimum)
            widget.set_row_heights([heights[row] for row in indices])
        widget.reset_undos()
        if session.filtered:
            self.status.set(f'{session.sheet_name} / フィルター本文 {len(session.visible_rows())}行 / 保存・CSV出力は非表示行を含む全行')

    def filtered_cell_reason(self, cell):
        if self.session.filtered and cell.row not in self.session.visible_rows():
            return 'フィルター中の見出しは編集・貼付けできません。先に解除してください'
        return ''

    def require_unfiltered_xlsx(self, *, all_sheets=False):
        if self.session and (self.session.filters if all_sheets else self.session.filtered):
            dialogs.showinfo('フィルター中', 'この操作の前にフィルターを解除してください。検索・置換はブック内の全シートで解除してください。', parent=self.root)
            return False
        return True

    def inline_coordinates(self, row, column):
        if not self.session.filtered:
            return row, column
        indices = self.session.displayed_indices()
        if self.sheet_widget.display_rows() != indices or not 0 <= row < len(indices):
            raise ValueError('表示条件が変わりました。セルを選び直してください')
        target = indices[row]
        if target + 1 not in self.session.visible_rows():
            raise ValueError('フィルター中の見出しは編集できません')
        return target, column

    @settled
    def xlsx_tool_action(self, action):
        if self.session is None:
            return
        try:
            action()
            self.search_tracker = XlsxSearch()
            self.render()
            self.sheet_widget.MT.focus_set()
        except Exception as error:
            record_failure('xlsx_sort_filter', error, warning=True)
            dialogs.showerror('XLSX操作中止', str(error), parent=self.root)

    def sort_xlsx(self):
        column, header, descending = self.sort_column.get(), self.sort_header.get(), self.sort_descending.get()
        self.xlsx_tool_action(lambda: self.session.sort_rows(int(column), header=header, descending=descending))

    def filter_xlsx(self):
        column, operator, query, header = self.sort_column.get(), self.filter_operator.get(), self.filter_query.get(), self.sort_header.get()
        self.xlsx_tool_action(lambda: self.session.set_filter(int(column),
            operator={'部分一致': 'contains', '完全一致': 'equals', '空': 'empty', '空以外': 'nonempty'}[operator],
            query=query, header=header))

    def clear_xlsx_filter(self):
        self.xlsx_tool_action(lambda: self.session.clear_filter())
