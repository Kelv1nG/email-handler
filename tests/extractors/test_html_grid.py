import pytest

from email_handler.extractors._html_grid import parse_html_grids


def test_parses_non_utf8_bytes_using_declared_charset():
    html = (
        b'<meta charset="windows-1252">'
        b"<table><tr><th>Name</th></tr><tr><td>Caf\xe9</td></tr></table>"
    )
    grids = parse_html_grids(html)
    assert grids[0].rows[1].values == ("Caf\u00e9",)


def test_parses_http_equiv_charset_declaration():
    html = (
        b'<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">'
        b"<table><tr><th>Name</th></tr><tr><td>Caf\xe9</td></tr></table>"
    )
    assert parse_html_grids(html)[0].rows[1].values == ("Caf\u00e9",)


def test_nested_table_is_isolated_and_following_sibling_keeps_dom_index():
    html = """
    <table>
      <tr><td>Before<table><caption>Child caption</caption><tr><th>Inner</th></tr><tr><td>Value</td></tr></table>After</td></tr>
      <tr><td>Tail</td></tr>
    </table>
    <table><tr><th>Sibling</th></tr></table>
    """
    grids = parse_html_grids(html)
    assert [grid.source_index for grid in grids] == [0, 1, 2]
    assert [row.values for row in grids[0].rows] == [("Before After",), ("Tail",)]
    assert grids[1].caption == "Child caption"
    assert grids[1].rows[0].values == ("Inner",)
    assert grids[1].rows[1].values == ("Value",)
    assert grids[2].rows[0].values == ("Sibling",)


def test_empty_table_is_skipped_but_still_advances_absolute_source_index():
    html = "<table></table><table><tr><th>A</th></tr></table>"
    grids = parse_html_grids(html)
    assert [grid.source_index for grid in grids] == [1]


def test_expands_rowspan_and_colspan_into_rectangular_rows():
    html = """
    <table>
      <tr><th rowspan="2">Account</th><th colspan="2">Amount</th></tr>
      <tr><th>Net</th><th>Tax</th></tr>
      <tr><td>A-1</td><td>10</td><td>2</td></tr>
    </table>
    """
    grid = parse_html_grids(html)[0]
    assert [row.values for row in grid.rows] == [
        ("Account", "Amount", "Amount"),
        ("Account", "Net", "Tax"),
        ("A-1", "10", "2"),
    ]


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("colspan", "0"),
        ("colspan", "-2"),
        ("colspan", "not-a-number"),
        ("rowspan", "0"),
        ("rowspan", "-2"),
        ("rowspan", "not-a-number"),
    ],
)
def test_invalid_spans_are_treated_as_one(name, value):
    html = f'<table><tr><td {name}="{value}">A</td><td>B</td></tr></table>'
    assert parse_html_grids(html)[0].rows[0].values == ("A", "B")


def test_active_rowspan_forces_new_cells_into_next_free_columns():
    html = """
    <table>
      <tr><td rowspan="2">A</td><td>B</td></tr>
      <tr><td colspan="2">C</td></tr>
    </table>
    """
    grid = parse_html_grids(html)[0]
    assert grid.rows[1].values == ("A", "C", "C")


def test_malformed_colspan_moves_to_the_next_contiguous_free_run():
    html = """
    <table>
      <tr><td>Lead</td><td rowspan="2">Held</td></tr>
      <tr><td colspan="2">Wide</td></tr>
    </table>
    """
    grid = parse_html_grids(html)[0]
    assert grid.rows[1].values == (None, "Held", "Wide", "Wide")


def test_ragged_rows_use_none_but_real_empty_cells_use_empty_string():
    html = """
    <table>
      <tr><th>A</th><th>B</th><th>C</th></tr>
      <tr><td>1</td><td></td></tr>
    </table>
    """
    grid = parse_html_grids(html)[0]
    assert grid.rows[1].values == ("1", "")


def test_ignores_script_style_and_nested_table_text():
    html = """
    <table><tr><td>
      Visible<script>hidden()</script><style>.hidden{}</style>
      <table><tr><td>nested</td></tr></table>
    </td></tr></table>
    """
    assert parse_html_grids(html)[0].rows[0].values == ("Visible",)
