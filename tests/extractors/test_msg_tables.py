import builtins
from pathlib import Path
import subprocess
import sys
from typing import Any, get_type_hints

import pytest

from email_handler.exceptions import MessageFileReadError
from email_handler.extractors.msg_tables import (
    extract_tables_from_msg,
    extracted_table_to_dataframe,
)
from email_handler.schemas.table import ExtractedTable, TableSelector


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


def test_extracted_table_to_dataframe_preserves_values_and_nulls():
    table = ExtractedTable(
        source_index=0,
        columns=["Account", "Amount"],
        rows=[["A-100", "125.50"], ["A-200", None]],
    )

    frame = extracted_table_to_dataframe(table)

    assert frame.columns == ["Account", "Amount"]
    assert frame.rows() == [("A-100", "125.50"), ("A-200", None)]


def test_extracted_table_to_dataframe_makes_duplicate_names_collision_safe():
    table = ExtractedTable(
        source_index=0,
        columns=["colA", "colA", "colA_2", "colB", "colB"],
        rows=[["1", "2", "3", "4", "5"]],
    )

    frame = extracted_table_to_dataframe(table)

    assert frame.columns == ["colA", "colA_3", "colA_2", "colB", "colB_2"]
    assert frame.rows() == [("1", "2", "3", "4", "5")]


def test_extracted_table_to_dataframe_explains_how_to_install_polars(monkeypatch):
    real_import = builtins.__import__

    def import_without_polars(name, *args, **kwargs):
        if name == "polars":
            raise ModuleNotFoundError("No module named 'polars'", name="polars")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", import_without_polars)
    table = ExtractedTable(source_index=0, columns=["A"], rows=[["1"]])

    with pytest.raises(
        ModuleNotFoundError, match=r"email-extractor\[polars\]"
    ) as raised:
        extracted_table_to_dataframe(table)

    assert isinstance(raised.value.__cause__, ModuleNotFoundError)


def test_extracted_table_to_dataframe_keeps_empty_columns_as_strings():
    table = ExtractedTable(
        source_index=0,
        columns=["Account", "Amount"],
        rows=[],
    )

    frame = extracted_table_to_dataframe(table)

    assert frame.shape == (0, 2)
    assert frame.columns == ["Account", "Amount"]
    assert [str(data_type) for data_type in frame.dtypes] == ["String", "String"]


def test_extracted_table_to_dataframe_runtime_annotations_resolve():
    assert get_type_hints(extracted_table_to_dataframe) == {
        "table": ExtractedTable,
        "return": Any,
    }


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


def test_public_dataframe_converter_import_does_not_require_polars():
    code = """
import builtins

real_import = builtins.__import__

def import_without_polars(name, *args, **kwargs):
    if name == 'polars':
        raise ModuleNotFoundError("No module named 'polars'", name='polars')
    return real_import(name, *args, **kwargs)

builtins.__import__ = import_without_polars
from email_handler.extractors import extracted_table_to_dataframe
assert callable(extracted_table_to_dataframe)
"""
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
