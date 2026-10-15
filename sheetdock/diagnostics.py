"""セル内容・パス・例外本文を含めない診断ログ。"""

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

_logger = None

OPERATIONS = {'startup', 'gui_callback', 'confirm', 'read', 'save_validate', 'save', 'edit', 'operation', 'settings_read', 'widths_save', 'settings_save', 'recent_save', 'geometry_save', 'search', 'header_toggle', 'reload', 'dialog_geometry_save', 'external_change', 'large_file', 'large_table', 'icon'}
ERROR_TYPES = {'OSError', 'PermissionError', 'FileNotFoundError', 'UnicodeDecodeError', 'UnicodeEncodeError', 'ValueError', 'TypeError', 'IndexError', 'RuntimeError', 'ExternalChangeError', 'TclError', 'JSONDecodeError'}


class QuietRotatingHandler(RotatingFileHandler):
    def handleError(self, record):
        # ログ書込失敗で編集処理を止めたり、標準エラーへ情報を出したりしない。
        pass


def log_directory():
    from .settings import default_path
    return default_path().parent / 'logs'


def configure_logging(directory=None, *, process_id=None, max_bytes=512 * 1024, backup_count=2):
    global _logger
    shutdown_logging()
    folder = Path(directory) if directory is not None else log_directory()
    pid = os.getpid() if process_id is None else int(process_id)
    path = folder / f'sheetdock-{pid}.log'
    try:
        folder.mkdir(parents=True, exist_ok=True)
        handler = QuietRotatingHandler(path, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8', delay=False)
    except OSError:
        return None
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    logger = logging.Logger('sheetdock', level=logging.WARNING)
    logger.propagate = False
    logger.addHandler(handler)
    _logger = logger
    return path


def record_failure(operation, exc=None, *, warning=False):
    if _logger is None:
        return
    operation = operation if operation in OPERATIONS else 'operation'
    name = type(exc).__name__ if exc is not None else '-'
    name = name if name in ERROR_TYPES or name == '-' else 'UnknownError'
    errno = exc.errno if isinstance(exc, OSError) and type(exc.errno) is int else '-'
    _logger.log(logging.WARNING if warning else logging.ERROR, 'operation=%s error=%s errno=%s', operation, name, errno)


def shutdown_logging():
    global _logger
    if _logger is not None:
        for handler in _logger.handlers[:]:
            handler.close()
            _logger.removeHandler(handler)
    _logger = None
