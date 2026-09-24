from pathlib import Path
import subprocess
import sys

import pytest

from email_handler.exceptions import TableExtractionError
import email_handler.extractors._html_grid as html_grid
from email_handler.extractors.html_tables import extract_html_tables
from email_handler.schemas.table import TableSelector


HTML = """
<table><tr><td>layout</td></tr></table>
<table>
  <caption>Balances</caption>
  <thead><tr><th colspan="2">Report</th></tr><tr><th> Account </th><th>Amount</th></tr></thead>
  <tbody><tr><td>A-1</td><td>10</td></tr></tbody>
</table>
<table><tr><td>account</td><td>amount</td></tr><tr><td>A-2</td><td>20</td></tr></table>
"""


def test_absolute_source_index_is_before_filtering():
    tables = extract_html_tables(HTML, TableSelector(source_index=1))
    assert [table.source_index for table in tables] == [1]
    assert tables[0].columns == ["Account", "Amount"]


def test_occurrence_is_after_case_insensitive_column_filtering():
    tables = extract_html_tables(
        HTML,
        TableSelector(required_columns=[" ACCOUNT ", "amount"], occurrence=1),
    )
    assert [table.source_index for table in tables] == [2]
    assert tables[0].rows == [["A-2", "20"]]


def test_omitted_occurrence_returns_every_column_match():
    tables = extract_html_tables(
        HTML,
        TableSelector(required_columns=["Account", "Amount"]),
    )
    assert [table.source_index for table in tables] == [1, 2]


def test_auto_header_prefers_last_nonempty_thead_row_and_caption():
    table = extract_html_tables(HTML, TableSelector(source_index=1))[0]
    assert table.caption == "Balances"
    assert table.columns == ["Account", "Amount"]
    assert table.rows == [["A-1", "10"]]


def test_explicit_header_discards_prior_rows_and_header():
    html = (
        "<table><tr><td>title</td></tr>"
        "<tr><td>A</td><td>B</td></tr>"
        "<tr><td>1</td><td>2</td></tr></table>"
    )
    table = extract_html_tables(html, TableSelector(header_row=1))[0]
    assert table.columns == ["A", "B"]
    assert table.rows == [["1", "2"]]


def test_explicit_header_accepts_a_structurally_present_blank_cell():
    html = "<table><tr><th></th></tr><tr><td>x</td></tr></table>"
    table = extract_html_tables(html, TableSelector(header_row=0))[0]
    assert table.columns == [""]
    assert table.rows == [["x"]]


def test_auto_header_uses_first_direct_th_row_without_thead():
    html = (
        "<table><tr><td>title</td></tr><tr><th>A</th></tr>"
        "<tr><td>1</td></tr></table>"
    )
    table = extract_html_tables(html)[0]
    assert table.columns == ["A"]
    assert table.rows == [["1"]]


def test_auto_header_falls_back_to_first_nonempty_row_without_th():
    html = (
        "<table><tr><td></td></tr><tr><td>A</td></tr>"
        "<tr><td>1</td></tr></table>"
    )
    table = extract_html_tables(html)[0]
    assert table.columns == ["A"]
    assert table.rows == [["1"]]


def test_auto_header_uses_first_direct_th_even_when_its_label_is_blank():
    html = "<table><tr><th></th></tr><tr><td>x</td></tr></table>"
    table = extract_html_tables(html)[0]
    assert table.columns == [""]
    assert table.rows == [["x"]]


@pytest.mark.parametrize("html", ["", b"", "<p>none</p>", "<table></table>"])
def test_normal_empty_inputs_return_empty_list(html):
    assert extract_html_tables(html) == []


def test_non_html_input_type_is_rejected():
    with pytest.raises(TypeError, match="str or bytes"):
        extract_html_tables(["not", "html"])


def test_str_and_equivalent_bytes_have_identical_results():
    html = (
        '<meta charset="utf-8"><table><tr><th>Caf\u00e9</th></tr>'
        "<tr><td>yes</td></tr></table>"
    )
    assert extract_html_tables(html) == extract_html_tables(html.encode("utf-8"))


def test_unexpected_grid_failure_is_chained(monkeypatch):
    def fail(_html):
        raise RuntimeError("parser exploded")

    monkeypatch.setattr("email_handler.extractors.html_tables.parse_html_grids", fail)
    with pytest.raises(TableExtractionError, match="parser exploded") as raised:
        extract_html_tables("<table></table>")
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_out_of_range_occurrence_is_empty():
    assert extract_html_tables(HTML, TableSelector(occurrence=99)) == []


def test_duplicate_and_blank_header_labels_are_preserved():
    html = (
        "<table><tr><th>A</th><th></th><th>A</th></tr>"
        "<tr><td>1</td><td></td><td>3</td></tr></table>"
    )
    table = extract_html_tables(html)[0]
    assert table.columns == ["A", "", "A"]
    assert table.rows == [["1", "", "3"]]


def test_pre_header_title_width_does_not_widen_selected_table():
    html = """
    <table>
      <tr><td colspan="4">Wide title</td></tr>
      <tr><th>A</th><th>B</th></tr>
      <tr><td>1</td><td>2</td></tr>
    </table>
    """
    table = extract_html_tables(html, TableSelector(header_row=1))[0]
    assert table.columns == ["A", "B"]
    assert table.rows == [["1", "2"]]


def test_public_result_pads_ragged_data_rows_with_none():
    html = """
    <table>
      <tr><th>A</th><th>B</th><th>C</th></tr>
      <tr><td>1</td><td></td></tr>
    </table>
    """
    assert extract_html_tables(html)[0].rows == [["1", "", None]]


def test_combined_rowspan_and_colspan_is_rectangular_in_public_result():
    html = """
    <table>
      <tr><th>One</th><th>Two</th><th>Three</th></tr>
      <tr><td rowspan="2">A</td><td>B</td><td>C</td></tr>
      <tr><td colspan="2">D</td></tr>
    </table>
    """
    table = extract_html_tables(html)[0]
    assert table.rows == [["A", "B", "C"], ["A", "D", "D"]]


def test_occurrence_counts_only_matches_but_source_index_stays_absolute():
    html = """
    <table></table>
    <table><tr><td>layout<table><tr><th>A</th></tr><tr><td>inner</td></tr></table></td></tr></table>
    <table><tr><th>A</th></tr><tr><td>sibling</td></tr></table>
    """
    table = extract_html_tables(
        html,
        TableSelector(required_columns=["A"], occurrence=1),
    )[0]
    assert table.source_index == 3
    assert table.rows == [["sibling"]]


def test_comments_do_not_contaminate_caption_headers_or_data():
    html = """
    <table>
      <caption>Balances<!-- internal note --></caption>
      <tr><th>Account<!-- heading note --></th></tr>
      <tr><td>A-1<!-- data note --></td></tr>
    </table>
    """

    tables = extract_html_tables(
        html,
        TableSelector(required_columns=["Account"]),
    )

    assert len(tables) == 1
    assert tables[0].caption == "Balances"
    assert tables[0].columns == ["Account"]
    assert tables[0].rows == [["A-1"]]


def test_source_index_skips_unselected_tables_before_grid_expansion(monkeypatch):
    original_expand_rows = html_grid._expand_rows

    def reject_unselected_table(table, rows):
        if table.get("id") == "unselected":
            raise AssertionError("unselected table was expanded")
        return original_expand_rows(table, rows)

    monkeypatch.setattr(html_grid, "_expand_rows", reject_unselected_table)
    html = (
        '<table id="selected"><tr><th>A</th></tr><tr><td>1</td></tr></table>'
        '<table id="unselected"><tr><th>B</th></tr></table>'
    )

    tables = extract_html_tables(html, TableSelector(source_index=0))

    assert [table.source_index for table in tables] == [0]
    assert tables[0].rows == [["1"]]


def test_oversized_colspan_fails_fast_with_chained_error():
    repository_root = Path(__file__).parents[2]
    script = """
from email_handler.exceptions import TableExtractionError
from email_handler.extractors import extract_html_tables

html = '<table><tr><th colspan="10001">A</th></tr></table>'
try:
    extract_html_tables(html)
except TableExtractionError as exc:
    assert isinstance(exc.__cause__, ValueError)
    assert "logical columns" in str(exc)
else:
    raise AssertionError("oversized colspan was accepted")
"""

    result = subprocess.run(
        [sys.executable, "-c", script],
        cwd=repository_root,
        check=False,
        capture_output=True,
        text=True,
        timeout=3,
    )

    assert result.returncode == 0, result.stderr
