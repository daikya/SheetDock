"""実行: python -m sheetdock.benchmark。合成CSVを一時フォルダーで測定。"""

import csv
import gc
import json
import platform
import tempfile
import time
import tracemalloc
from pathlib import Path

from .gui import Session
from .storage import load_checked


def main():
    results = []
    with tempfile.TemporaryDirectory() as folder:
        for count in (1000, 10_000, 50_000):
            path = Path(folder) / 'sample.csv'
            with path.open('w', encoding='utf-8-sig', newline='') as stream:
                writer = csv.writer(stream)
                writer.writerow([f'列{i}' for i in range(10)])
                for row in range(count):
                    writer.writerow([f'{row:08d}', *[f'日本語データ{row:08d}-{col}' for col in range(9)]])
            gc.collect()
            tracemalloc.start()
            start = time.perf_counter()
            document, stamp = load_checked(path, encoding='utf-8')
            session = Session(document, True)
            header, body = session.view()
            session.editor.set_cell(1, 0, '00999999')
            elapsed = time.perf_counter() - start
            _, peak = tracemalloc.get_traced_memory()
            tracemalloc.stop()
            result = {'rows': count + 1, 'columns': 10, 'bytes': path.stat().st_size, 'seconds': round(elapsed, 3), 'python_peak_mib': round(peak / 1024**2, 2)}
            results.append(result)
            print(json.dumps(result), flush=True)
            del document, session, header, body
    print(json.dumps({'python': platform.python_version(), 'platform': platform.system(), 'scope': 'load/model/view-copy/one-edit; no Tk rendering; tracemalloc enabled', 'results': results}, ensure_ascii=False))


if __name__ == '__main__':
    main()
