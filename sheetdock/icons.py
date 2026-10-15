"""ソース版・exe版の共通アイコン。"""

import sys
import tkinter as tk
from pathlib import Path

from .diagnostics import record_failure


def icon_path():
    return Path(__file__).resolve().parent / 'assets' / 'sheetdock.ico'


def apply_icon(root):
    try:
        if sys.platform == 'win32':
            root.iconbitmap(default=str(icon_path()))
        else:
            image = tk.PhotoImage(master=root, file=str(icon_path().with_suffix('.png')))
            root.iconphoto(True, image)
            root._sheetdock_icon = image
    except (tk.TclError, OSError) as exc:
        record_failure('icon', exc, warning=True)
