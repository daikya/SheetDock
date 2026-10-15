import errno

import pytest

import sheetdock.diagnostics as diagnostics


@pytest.fixture(autouse=True)
def reset_logging():
    diagnostics.shutdown_logging()
    yield
    diagnostics.shutdown_logging()


def test_default_directory_uses_localappdata(tmp_path, monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    assert diagnostics.log_directory() == tmp_path / 'SheetDock' / 'logs'


def test_error_logs_only_operation_type_and_errno(tmp_path):
    path = diagnostics.configure_logging(tmp_path, process_id=123)
    diagnostics.record_failure('save', PermissionError(errno.EACCES, '秘密のセル内容', 'private-contact.csv'))
    text = path.read_text(encoding='utf-8')
    assert 'ERROR' in text
    assert 'operation=save' in text
    assert 'error=PermissionError' in text
    assert f'errno={errno.EACCES}' in text
    assert '秘密' not in text
    assert 'private-contact' not in text
    assert 'Traceback' not in text


def test_unicode_error_does_not_log_input_bytes(tmp_path):
    path = diagnostics.configure_logging(tmp_path)
    diagnostics.record_failure('read', UnicodeDecodeError('utf-8', b'PRIVATE-CELL\xff', 12, 13, 'PRIVATE-REASON'))
    text = path.read_text()
    assert 'UnicodeDecodeError' in text
    assert 'PRIVATE' not in text
    assert '\\xff' not in text


def test_unknown_operation_and_exception_name_are_sanitized(tmp_path):
    path = diagnostics.configure_logging(tmp_path)
    custom = type('PRIVATE_CLASS_NAME', (Exception,), {})
    diagnostics.record_failure('PRIVATE_QUERY', custom('PRIVATE_VALUE'))
    text = path.read_text()
    assert 'operation=operation error=UnknownError' in text
    assert 'PRIVATE' not in text


def test_warning_event_without_exception(tmp_path):
    path = diagnostics.configure_logging(tmp_path)
    diagnostics.record_failure('external_change', warning=True)
    text = path.read_text()
    assert 'WARNING' in text
    assert 'operation=external_change error=-' in text


def test_logs_for_processes_have_separate_paths(tmp_path):
    a = diagnostics.configure_logging(tmp_path, process_id=123)
    diagnostics.record_failure('read', ValueError())
    b = diagnostics.configure_logging(tmp_path, process_id=456)
    diagnostics.record_failure('save', OSError())
    assert a != b
    assert 'operation=save' not in a.read_text()
    assert 'operation=read' not in b.read_text()


def test_rotation_is_bounded_per_process(tmp_path):
    path = diagnostics.configure_logging(tmp_path, process_id=123, max_bytes=150, backup_count=2)
    for _ in range(20):
        diagnostics.record_failure('save', OSError())
    diagnostics.shutdown_logging()
    files = list(tmp_path.glob('sheetdock-123.log*'))
    assert len(files) == 3
    assert all(f.stat().st_size < 200 for f in files)


def test_unwritable_log_location_does_not_block_application(tmp_path):
    path = tmp_path / 'not-a-directory'
    path.write_text('occupied')
    assert diagnostics.configure_logging(path) is None
    diagnostics.record_failure('save', OSError('private'))


def test_log_write_failure_does_not_raise(tmp_path):
    diagnostics.configure_logging(tmp_path)
    handler = diagnostics._logger.handlers[0]
    handler.stream.close()
    class BrokenStream:
        def write(self, text):
            raise OSError('private storage failure')
        def tell(self):
            return 0
        def seek(self, *args):
            return 0
        def flush(self):
            pass
        def close(self):
            pass
    handler.stream = BrokenStream()
    diagnostics.record_failure('save', OSError())


def test_no_configuration_does_not_create_log_files(tmp_path, monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    diagnostics.record_failure('read', ValueError())
    assert not (tmp_path / 'SheetDock').exists()
