"""Validated inline editing and Tk editor lifecycle for XLSX panes."""
from functools import wraps
import tkinter as tk

from . import dialogs
from .xlsx_display import format_value
from .xlsx_extent import extent
from .xlsx_axis import VIEW_ROWS, VIEW_COLUMNS


class InlineBridge:
    """Validate a single widget edit through the authoritative workbook session."""
    def __init__(self, session, refresh, notice):
        self.session, self.refresh, self.notice = session, refresh, notice
        self.context = None
        self.last_error = None
        self.pending_error = None

    def begin(self, event):
        from .xlsx_gui import protection_reason, cell_state, display_value
        if self.pending_error:
            self.notice('入力エラーを確認してから再度編集してください')
            return None
        self.context, self.last_error = None, None
        session = self.session
        row, column = event['row'], event['column']
        if session is None:
            return None
        rows, columns = extent(session.sheet)
        if not (0 <= row < min(rows, VIEW_ROWS) and 0 <= column < min(columns, VIEW_COLUMNS)):
            self.notice('表示範囲内のセルを選択してください')
            return None
        cell = session.sheet.cell(row + 1, column + 1)
        reason = protection_reason(session.probe, session.sheet, cell)
        if reason:
            self.notice(reason)
            return None
        self.context = (session, session.probe.book, session.sheet_name, row, column, cell_state(cell))
        # Double click/F2/Return starts with raw data, not formatted display text.
        # Character typing and BackSpace retain the widget's replace-input behavior.
        return display_value(cell.value) if event.get('key', '??') in {'??', 'F2', 'Return'} else event['value']

    def validate(self, event):
        from .xlsx_gui import cell_state
        context, self.context = self.context, None
        try:
            if context is None:
                raise ValueError('編集開始情報がありません。セルを選び直してください')
            session, book, name, row, column, before = context
            if (session is not self.session or book is not session.probe.book or name != session.sheet_name
                    or (row, column) != (event['row'], event['column'])):
                raise ValueError('編集中に対象が変わりました。再度入力してください')
            cell = session.sheet.cell(row + 1, column + 1)
            if cell_state(cell) != before:
                raise ValueError('編集中にセルが変わりました。再度入力してください')
            if not isinstance(event['value'], str):
                raise TypeError('文字列の入力を指定してください')
            session.edit(row, column, event['value'])
            self.last_error = None
            value = format_value(cell.value, cell.number_format)[0]
        except Exception as exc:
            self.last_error = self.pending_error = str(exc)
            value = None  # Reject the widget mutation too.
        self.refresh()
        return value

    def cancel(self):
        self.context, self.last_error, self.pending_error = None, None, None


def settled(method):
    """Only run a pane action after pending input has passed validation."""
    @wraps(method)
    def guarded(self, *args, **kwargs):
        if not self.finish_inline_edit():
            return 'break'
        return method(self, *args, **kwargs)
    return guarded


class InlineEditing:
    def setup_inline(self):
        self._inline_flush_id = None
        self._inline = InlineBridge(self.session, self.queue_inline_refresh, self.detail.set)
        sheet = self.sheet_widget
        sheet.extra_bindings('begin_edit_cell', self.begin_inline_edit)
        sheet.edit_validation(self._inline.validate)
        sheet.bind_key_text_editor('<Escape>', self.cancel_inline_edit)
        sheet.bind_key_text_editor('<FocusOut>', self.inline_focus_out)
        for key, redo in [('<Control-z>', False), ('<Control-Z>', False),
                          ('<Control-Shift-z>', True), ('<Control-Shift-Z>', True)]:
            sheet.bind_key_text_editor(key, lambda event, redo=redo: self.inline_text_history(event, redo))

    def begin_inline_edit(self, event):
        if self._rendering:
            return None
        self._inline.session = self.session
        return self._inline.begin(event)

    def inline_text_history(self, event, redo=False):
        try:
            event.widget.edit_redo() if redo else event.widget.edit_undo()
        except tk.TclError:
            pass
        return 'break'

    def cancel_inline_edit(self, event=None):
        self._inline.cancel()
        self.sheet_widget.close_text_editor(set_data=False)
        return 'break'

    def inline_focus_out(self, event):
        if event.widget.focus_get() is event.widget:
            return 'break'
        event.keysym = 'FocusOut'
        self.sheet_widget.MT.close_text_editor(event)
        return 'break'

    def queue_inline_refresh(self):
        if self._inline_flush_id is None:
            self._inline_flush_id = self.root.after_idle(self.flush_inline_edit)

    def flush_inline_edit(self):
        if self._inline_flush_id is not None:
            self.root.after_cancel(self._inline_flush_id)
            self._inline_flush_id = None
        error, self._inline.pending_error = self._inline.pending_error, None
        # Native validation has already updated the formatted cell display.
        # Refresh controls only; a new editor may already be open.
        self.selection_changed()
        if getattr(self, '_on_change', None):
            self._on_change()
        if error:
            self.detail.set('入力取消：' + error)
            dialogs.showerror('直接編集の入力エラー', error, parent=self.root)

    def finish_inline_edit(self):
        bridge = getattr(self, '_inline', None)
        if bridge is None:
            return True
        bridge.session = self.session
        editor = self.sheet_widget.MT.text_editor
        was_open, pending_error = editor.open, bridge.pending_error
        if was_open:
            # A notebook may already have hidden this Text widget. The native
            # close method refuses to commit when focus is None, so settle the
            # same validated model transaction without depending on Tk focus.
            row, column = editor.coords
            display = bridge.validate({'row': row, 'column': column,
                                       'value': self.sheet_widget.get_text_editor_value()})
            self.sheet_widget.MT.hide_text_editor_and_dropdown()
            if display is not None:
                self.sheet_widget.set_cell_data(row, column, value=display, redraw=True)
        if self._inline_flush_id is not None:
            self.flush_inline_edit()
        return not (pending_error or was_open and bridge.last_error)

