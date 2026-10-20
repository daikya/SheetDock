"""Lightweight workbook sheet tabs sharing a single table in the parent pane."""
import tkinter as tk
from tkinter import ttk


class SheetTabs(ttk.Frame):
    def __init__(self, parent, on_select):
        super().__init__(parent)
        self.on_select = on_select
        self.pages = {}
        self.current = None
        self.columnconfigure(0, weight=1)
        # Pages contain no tables: only the active worksheet is rendered.
        self.book = ttk.Notebook(self, height=0)
        self.book.grid(row=0, column=0, sticky='ew')
        self.book.bind('<<NotebookTabChanged>>', self.changed)
        menu_button = ttk.Menubutton(self, text='シート一覧')
        menu_button.grid(row=0, column=1, padx=4)
        self.menu = tk.Menu(menu_button, tearoff=False, postcommand=self.fill_menu)
        menu_button.configure(menu=self.menu)

    def sync(self, names, current):
        names = tuple(names)
        if current not in names:
            raise ValueError('選択シートが存在しません')
        self.current = current
        if tuple(self.pages) != names:
            for page in self.pages.values():
                self.book.forget(page)
                page.destroy()
            self.pages = {}
            for name in names:
                page = ttk.Frame(self.book, height=0)
                self.book.add(page, text=name)
                self.pages[name] = page
        self.book.select(self.pages[current])

    def changed(self, event=None):
        selected = self.book.select()
        name = next((name for name, page in self.pages.items()
                     if str(page) == selected), None)
        if name is not None and name != self.current:
            self.current = name
            self.on_select(name)

    def choose(self, name):
        if name in self.pages:
            self.book.select(self.pages[name])
            self.changed()

    def fill_menu(self):
        self.menu.delete(0, 'end')
        for name in self.pages:
            marker = '✓ ' if name == self.current else '   '
            self.menu.add_command(label=marker + name, command=lambda n=name: self.choose(n))
