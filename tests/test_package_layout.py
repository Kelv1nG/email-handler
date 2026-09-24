"""Contracts for consumers importing the installed package."""

from pathlib import Path
import subprocess
import sys


def test_namespaced_public_imports_work_without_outlook():
    repository_root = Path(__file__).parents[1]
    script = """
import sys
from email_handler import exceptions
from email_handler.extractors import extract_html_tables, extract_tables_from_msg
from email_handler.protocols import EmailProvider, TableExtractingEmailProvider
from email_handler.schemas import ExtractedTable, TableSelector

table = extract_html_tables(
    '<table><tr><th>Account</th></tr><tr><td>A-1</td></tr></table>',
    TableSelector(required_columns=['Account']),
)[0]
assert isinstance(table, ExtractedTable)
assert table.rows == [['A-1']]
assert callable(extract_tables_from_msg)
assert exceptions.TableExtractionError.__module__ == 'email_handler.exceptions'
assert 'win32com' not in sys.modules
assert 'email_handler.providers.outlook' not in sys.modules
"""
    result = subprocess.run(
        [sys.executable, "-c", script],
        cwd=repository_root,
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
