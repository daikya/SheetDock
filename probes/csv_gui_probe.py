"""実行: python probes/csv_gui_probe.py。単一CSVのGUI接続検証。"""

import copy
import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tksheet import Sheet
from csv_probe import CsvDocument, load_csv, save_csv


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


class CsvGuiProbe:
    def __init__(self, root):
        self.root = root
        self.document = None
        self.header = None
        self.source = None
        self.last_saved_rows = None
        root.title('SheetDock — CSV GUI接続検証')
        root.geometry('1100x700')
        root.protocol('WM_DELETE_WINDOW', self.close)
        self.encoding = tk.StringVar(value='utf-8')
        self.delimiter = tk.StringVar(value='カンマ')
        self.use_header = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value='読込条件を選んでCSVを開いてください。保存は別名のみです。')
        controls = ttk.Frame(root)
        controls.pack(fill='x', padx=10, pady=10)
        ttk.Label(controls, text='次に開くCSV：文字コード').pack(side='left')
        ttk.Combobox(controls, textvariable=self.encoding, values=('utf-8', 'cp932'), state='readonly', width=10).pack(side='left', padx=5)
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
            messagebox.showerror('別名保存', 'この検証では読込元と異なる保存先を選んでください。', parent=self.root)
            return
        try:
            save_csv(edited, target, backup=True)
            # 別名出力なので読込元は変えず、保存済みの編集内容だけ更新する。
            self.last_saved_rows = copy.deepcopy(edited.rows)
        except Exception as exc:
            messagebox.showerror('保存失敗', str(exc), parent=self.root)
            return
        self.status.set(f'保存・再読込検証成功：{target}')
        messagebox.showinfo('保存完了', '別名保存と再読込での一致確認が完了しました。', parent=self.root)

    def close(self):
        if self.may_discard():
            self.root.destroy()


def main():
    root = tk.Tk()
    CsvGuiProbe(root)
    root.mainloop()


if __name__ == '__main__':
    main()
