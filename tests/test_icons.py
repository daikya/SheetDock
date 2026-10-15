import struct
from types import SimpleNamespace

import sheetdock.icons as icons


def test_icon_contains_multiple_windows_sizes():
    data = icons.icon_path().read_bytes()
    reserved, kind, count = struct.unpack_from('<HHH', data)
    assert (reserved, kind) == (0, 1)
    sizes = {((data[6 + i * 16] or 256), (data[7 + i * 16] or 256)) for i in range(count)}
    assert {(16, 16), (32, 32), (48, 48), (256, 256)} <= sizes


def test_windows_default_icon_applies_to_future_windows(monkeypatch):
    calls = []
    root = SimpleNamespace(iconbitmap=lambda **kwargs: calls.append(kwargs))
    monkeypatch.setattr(icons.sys, 'platform', 'win32')
    icons.apply_icon(root)
    assert calls == [{'default': str(icons.icon_path())}]
