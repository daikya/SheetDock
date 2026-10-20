"""複数ファイルの状態と保存先の競合管理。GUIには依存しない。"""

import os
from dataclasses import dataclass
from pathlib import Path


def path_key(path):
    return os.path.normcase(str(Path(path).resolve()))


def session_dirty(session):
    return session.editor.dirty if hasattr(session, 'editor') else session.dirty


@dataclass(eq=False)
class DocumentTab:
    source: Path
    session: object
    stamp: object = None

    @property
    def title(self):
        return self.source.name + (' *' if session_dirty(self.session) else '')


class Documents:
    def __init__(self):
        self.tabs = []

    def find(self, path):
        for tab in self.tabs:
            if path_key(tab.source) == path_key(path):
                return tab
            if tab.source.exists() and Path(path).exists() and os.path.samefile(tab.source, path):
                return tab
        return None

    def add(self, path, session, stamp=None):
        existing = self.find(path)
        if existing is not None:
            return existing
        tab = DocumentTab(Path(path), session, stamp)
        self.tabs.append(tab)
        return tab

    def check_save_as(self, tab, path):
        existing = self.find(path)
        if existing is tab:
            raise ValueError('別名保存では現在のファイルと異なる保存先を選んでください。')
        if existing is not None:
            raise ValueError('保存先のファイルは別のタブで開かれています。そのタブを閉じてから保存してください。')

    def saved_as(self, tab, path):
        # 書込成功後にだけ、パスと保存済み状態を更新する。
        self.check_save_as(tab, path)
        tab.source = Path(path)
        tab.session.editor.mark_saved()

    def remove(self, tab):
        self.tabs.remove(tab)

    def may_close(self, tab, decide, save):
        if not session_dirty(tab.session):
            return True
        choice = decide(tab)
        if choice == 'discard':
            return True
        if choice == 'save':
            return bool(save(tab)) and not session_dirty(tab.session)
        return False
