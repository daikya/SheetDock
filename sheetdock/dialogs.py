"""親中央への初期配置と、サブ画面の前回配置を共通管理する。"""

import hashlib
import tkinter as tk
from tkinter import ttk

from .settings import SettingsStore, geometry_parts, visible_geometry
from .diagnostics import record_failure


def placement(saved, parent_rect, requested_size, screen_size):
    restored = visible_geometry(saved, *screen_size, minimum=(80, 50))
    if restored:
        return restored
    x, y, w, h = parent_rect
    dw, dh = requested_size
    dw, dh = max(80, dw), max(50, dh)
    proposed = f'{dw}x{dh}{x + (w - dw) // 2:+d}{y + (h - dh) // 2:+d}'
    return visible_geometry(proposed, *screen_size, minimum=(80, 50))


def show_placed(window, parent, key):
    store = SettingsStore()
    window.update_idletasks()
    screen = parent.winfo_screenwidth(), parent.winfo_screenheight()
    if parent.winfo_ismapped():
        parent_rect = parent.winfo_x(), parent.winfo_y(), parent.winfo_width(), parent.winfo_height()
    else:
        parent_rect = 0, 0, *screen
    geometry = placement(store.load().dialog_geometry.get(key), parent_rect, (window.winfo_reqwidth(), window.winfo_reqheight()), screen)
    remembered = [geometry]
    persisted = [False]
    def configure(event):
        if event.widget is window and window.state() == 'normal':
            value = window.geometry()
            if geometry_parts(value, (80, 50)):
                remembered[0] = value
    def persist(event):
        if event.widget is window and not persisted[0]:
            persisted[0] = True
            try:
                store.update(dialog_key=key, dialog_geometry=remembered[0])
            except (OSError, RuntimeError, ValueError) as exc:
                record_failure('dialog_geometry_save', exc, warning=True)
    window.geometry(geometry)
    window.deiconify()
    window.update_idletasks()
    window.geometry(geometry)
    window.bind('<Configure>', configure, add='+')
    window.bind('<Destroy>', persist, add='+')


def dialog_key(kind, title):
    return kind + ':' + hashlib.sha256(title.encode('utf-8')).hexdigest()[:16]


def _window(parent, title):
    window = tk.Toplevel(parent)
    window.withdraw()
    window.title(title)
    window.transient(parent)
    return window


def _message(title, message, *, parent, choices):
    window = _window(parent, title)
    result = [None]
    ttk.Label(window, text=message, wraplength=520, padding=16).pack(fill='both', expand=True)
    buttons = ttk.Frame(window, padding=10)
    buttons.pack(fill='x')
    def finish(value):
        result[0] = value
        window.destroy()
    first = None
    for label, value in choices:
        button = ttk.Button(buttons, text=label, command=lambda v=value: finish(v))
        button.pack(side='left', padx=6)
        if first is None:
            first = button
    window.bind('<Escape>', lambda event: finish(None))
    window.bind('<Return>', lambda event: finish(choices[0][1]))
    show_placed(window, parent, dialog_key('message', title))
    window.grab_set()
    first.focus_set()
    parent.wait_window(window)
    return result[0]


def showinfo(title, message, *, parent):
    return _message(title, message, parent=parent, choices=[('OK', 'ok')])


showerror = showinfo
showwarning = showinfo


def askyesnocancel(title, message, *, parent):
    return _message(title, message, parent=parent, choices=[('はい', True), ('いいえ', False), ('キャンセル', None)])


def askyesno(title, message, *, parent):
    return _message(title, message, parent=parent, choices=[('はい', True), ('いいえ', False)]) is True


def _input(title, prompt, *, parent, initialvalue='', integer=False, minvalue=None, maxvalue=None):
    window = _window(parent, title)
    result = [None]
    ttk.Label(window, text=prompt, wraplength=480, padding=12).pack(fill='x')
    value = tk.StringVar(value=str(initialvalue) if initialvalue is not None else '')
    entry = ttk.Entry(window, textvariable=value, width=40)
    entry.pack(fill='x', padx=12)
    error = tk.StringVar(value='')
    ttk.Label(window, textvariable=error).pack(fill='x', padx=12, pady=5)
    def accept():
        output = value.get()
        if integer:
            try:
                output = int(output)
            except ValueError:
                error.set('整数を入力してください。')
                return
            if (minvalue is not None and output < minvalue) or (maxvalue is not None and output > maxvalue):
                error.set('指定された範囲内の整数を入力してください。')
                return
        result[0] = output
        window.destroy()
    buttons = ttk.Frame(window, padding=10)
    buttons.pack(fill='x')
    ttk.Button(buttons, text='OK', command=accept).pack(side='left', padx=5)
    ttk.Button(buttons, text='キャンセル', command=window.destroy).pack(side='left', padx=5)
    window.bind('<Return>', lambda event: accept())
    window.bind('<Escape>', lambda event: window.destroy())
    show_placed(window, parent, dialog_key('input', title))
    window.grab_set()
    entry.focus_set()
    entry.selection_range(0, 'end')
    parent.wait_window(window)
    return result[0]


def askstring(title, prompt, **kwargs):
    return _input(title, prompt, **kwargs)


def askinteger(title, prompt, **kwargs):
    return _input(title, prompt, integer=True, **kwargs)
