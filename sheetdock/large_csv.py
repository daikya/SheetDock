"""大容量警告の基準。読込禁止の制限にはしない。"""

from pathlib import Path

FILE_WARNING_BYTES = 5 * 1024 * 1024
CELL_WARNING_COUNT = 100_000


def large_file(path):
    return Path(path).stat().st_size >= FILE_WARNING_BYTES


def table_size(rows):
    return len(rows), sum(len(row) for row in rows)


def large_table(rows):
    return table_size(rows)[1] >= CELL_WARNING_COUNT
