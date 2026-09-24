from pathlib import Path
import subprocess
import sys

import pytest

from email_handler.exceptions import MessageFileReadError
from email_handler.extractors.msg_tables import extract_tables_from_msg
from email_handler.schemas.table import TableSelector


class FakeMessage:
    def __init__(self, html_body=None, body_failure=None):
        self._html_body = html_body
        self.body_failure = body_failure
        self.closed = False

    @property
    def htmlBody(self):
        if self.body_failure:
            raise self.body_failure
        return self._html_body

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.close()

    def close(self):
        self.closed = True


def test_passes_html_bytes_to_shared_parser(tmp_path, monkeypatch):
    path = tmp_path / "report.msg"
    path.write_bytes(b"placeholder")
    message = FakeMessage(
        b"<table><tr><th>A</th></tr><tr><td>1</td></tr></table>"
    )
    monkeypatch.setattr(
        "email_handler.extractors.msg_tables.extract_msg.openMsg", lambda _path: message
    )

    tables = extract_tables_from_msg(
        path, TableSelector(required_columns=["A"])
    )

    assert tables[0].rows == [["1"]]
    assert message.closed is True


def test_delegates_the_exact_msg_html_bytes_and_selector(tmp_path, monkeypatch):
    path = tmp_path / "raw.msg"
    path.write_bytes(b"placeholder")
    html_body = (
        b"<meta charset=windows-1252><table><tr><td>Caf\xe9</td></tr></table>"
    )
    message = FakeMessage(html_body)
    selector = TableSelector(source_index=0)
    monkeypatch.setattr(
        "email_handler.extractors.msg_tables.extract_msg.openMsg",
        lambda _path: message,
    )

    def capture(html, received_selector):
        assert html is html_body
        assert received_selector is selector
        return []

    monkeypatch.setattr("email_handler.extractors.msg_tables.extract_html_tables", capture)
    assert extract_tables_from_msg(path, selector) == []
    assert message.closed is True


def test_empty_msg_html_is_successful_empty_result(tmp_path, monkeypatch):
    path = tmp_path / "empty.msg"
    path.write_bytes(b"placeholder")
    monkeypatch.setattr(
        "email_handler.extractors.msg_tables.extract_msg.openMsg",
        lambda _path: FakeMessage(None),
    )
    assert extract_tables_from_msg(path) == []


def test_path_validation_order(tmp_path):
    missing = tmp_path / "missing.txt"
    with pytest.raises(FileNotFoundError):
        extract_tables_from_msg(missing)

    directory = tmp_path / "folder.msg"
    directory.mkdir()
    with pytest.raises(IsADirectoryError):
        extract_tables_from_msg(directory)

    wrong_suffix = tmp_path / "message.eml"
    wrong_suffix.write_bytes(b"x")
    with pytest.raises(ValueError, match=r"\.msg"):
        extract_tables_from_msg(wrong_suffix)


def test_unexpected_path_access_failure_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "permission.msg"

    def fail(_path):
        raise PermissionError("access denied")

    monkeypatch.setattr("email_handler.extractors.msg_tables._validated_msg_path", fail)
    with pytest.raises(MessageFileReadError, match="permission.msg") as raised:
        extract_tables_from_msg(path)
    assert isinstance(raised.value.__cause__, PermissionError)


def test_corrupt_msg_error_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "bad.msg"
    path.write_bytes(b"bad")

    def fail(_path):
        raise RuntimeError("bad compound file")

    monkeypatch.setattr("email_handler.extractors.msg_tables.extract_msg.openMsg", fail)
    with pytest.raises(MessageFileReadError, match="bad.msg") as raised:
        extract_tables_from_msg(path)
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_html_body_failure_closes_message_and_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "broken-body.msg"
    path.write_bytes(b"placeholder")
    message = FakeMessage(body_failure=RuntimeError("body unavailable"))
    monkeypatch.setattr(
        "email_handler.extractors.msg_tables.extract_msg.openMsg", lambda _path: message
    )

    with pytest.raises(MessageFileReadError, match="broken-body.msg") as raised:
        extract_tables_from_msg(path)

    assert message.closed is True
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_parser_failure_closes_message_and_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "parser-failure.msg"
    path.write_bytes(b"placeholder")
    message = FakeMessage(b"<table></table>")
    monkeypatch.setattr(
        "email_handler.extractors.msg_tables.extract_msg.openMsg", lambda _path: message
    )

    def fail(_html, _selector):
        raise RuntimeError("parser exploded")

    monkeypatch.setattr("email_handler.extractors.msg_tables.extract_html_tables", fail)
    with pytest.raises(MessageFileReadError, match="parser-failure.msg") as raised:
        extract_tables_from_msg(path)

    assert message.closed is True
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_msg_reader_import_does_not_import_outlook_dependencies():
    code = (
        "import sys; import email_handler.extractors.msg_tables; "
        "assert 'win32com' not in sys.modules; "
        "assert 'email_handler.providers.outlook' not in sys.modules"
    )
    subprocess.run([sys.executable, "-c", code], check=True)


def test_real_msg_fixture_extracts_two_tables_without_outlook():
    fixture = Path(__file__).parents[1] / "fixtures" / "table_message.msg"
    tables = extract_tables_from_msg(fixture)
    assert [table.columns for table in tables] == [
        ["Item", "Quantity"],
        ["Account", "Amount"],
    ]
    assert tables[0].rows == [["Caf\u00e9", "2"]]
    assert tables[1].rows == [["A-100", "125.50"]]
