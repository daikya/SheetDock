"""SheetDock tksheet操作検証。本体アプリではなく、ファイルは保存しない。

Windows / Python 3.12:
  python -m pip install tksheet==7.6.0
  python probes/tksheet_probe.py

確認: 日本語編集、範囲選択、Excel相互コピー（00123・タブ・セル内改行）、
複数行列の右クリック挿入/削除、Ctrl+Z / Ctrl+Shift+Z、検索置換のUndo。
合否は実際の結果を記録する。数値に変換された場合は不合格。
"""

import platform
import tkinter as tk
from importlib.metadata import version
from tkinter import ttk

from tksheet import Sheet


def main():
    root = tk.Tk()
    root.title("SheetDock — tksheet 技術検証（保存機能なし）")
    root.geometry("1000x650")
    info = f"Python {platform.python_version()} / tksheet {version('tksheet')} / Tk {root.tk.call('info', 'patchlevel')}"
    ttk.Label(root, text=info).pack(anchor="w", padx=10, pady=8)
    ttk.Label(root, text="右クリック：行列操作　Ctrl+C/V：コピー/貼付け　Ctrl+Z：Undo　Ctrl+Shift+Z：Redo").pack(anchor="w", padx=10)
    sheet = Sheet(root, data=[
        ["00123", "日本語", "カンマ,あり", '引用符"あり'],
        ["00000", "改行\n二行目", "タブ\tあり", ""],
        ["1E10", "2026-10-13", "=SUM(A1:A2)", "CSVでは文字列"],
        ["末尾", "", "", ""],
    ])
    sheet.pack(fill="both", expand=True, padx=10, pady=10)
    sheet.enable_bindings("all")
    bar = ttk.Frame(root)
    bar.pack(fill="x", padx=10, pady=8)
    ttk.Button(bar, text="Undo", command=sheet.undo).pack(side="left")
    ttk.Button(bar, text="Redo", command=sheet.redo).pack(side="left")
    ttk.Button(bar, text="検索・置換", command=lambda: sheet.open_find(focus=True)).pack(side="left")
    status = tk.StringVar(value="検証用データのみ。CSVはすべて文字列として維持する必要があります。")
    def check_strings():
        values = sheet.get_sheet_data()
        non_strings = sum(not isinstance(value, str) for row in values for value in row)
        status.set(f"行数={len(values)} / 文字列以外={non_strings}（0が合格）")
    ttk.Button(bar, text="データ型を確認", command=check_strings).pack(side="left")
    ttk.Label(root, textvariable=status).pack(anchor="w", padx=10, pady=8)
    root.mainloop()


if __name__ == "__main__":
    main()
