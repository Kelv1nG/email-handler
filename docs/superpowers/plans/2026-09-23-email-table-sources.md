# Email Table Sources Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (\`- [ ]\`) syntax for tracking.

**Goal:** Build backward-compatible Pydantic table extraction from live Outlook messages and standalone MSG files, plus a safe utility for saving cached live messages as Unicode MSG files and runnable examples for all three workflows.

**Architecture:** A pure HTML layer accepts str or bytes and converts every HTML table into a normalized grid before a selector facade chooses headers, columns, and occurrences. Source adapters remain explicit: OutlookProvider supplies live HTML strings, extract_tables_from_msg supplies MSG HTML bytes, and OutlookProvider.save_message persists cached MailItems; all provider-specific code remains outside the parser.

**Tech Stack:** Python 3.13+, Pydantic 2, Beautiful Soup 4 with html.parser, extract-msg 0.56.x, pywin32 on Windows, pytest, uv.

**Spec:** docs/superpowers/specs/2026-09-23-email-table-sources-design.md

## Global Constraints

- Preserve the exact fields and serialized shapes of EmailRecord, SearchQuery, QueryResult, and ExtractionResult.
- Do not add extract_tables or save_message to protocols/provider.py::EmailProvider.
- Keep EmailRecord.body sourced from Outlook message.Body; read HTMLBody only inside the new live extraction method.
- The standalone MSG reader must not import win32com, providers.outlook, or require Outlook.
- Accept both str and bytes at the pure HTML boundary; never hard-code UTF-8 decoding for MSG HTML bytes.
- Use zero-based absolute source_index and zero-based occurrence after required-column filtering.
- Keep the canonical table payload as columns: list[str] and rows: list[list[str | None]]; do not add pandas.
- Use Beautiful Soup with html.parser; do not add lxml, html5lib, a browser engine, or CSS-layout reconstruction.
- Use extract-msg>=0.56,<0.57 and condition pywin32>=311 on sys_platform == "win32".
- Saving must use Outlook olMSGUnicode value 9, refuse any destination that exists at preflight, and never send, mutate, move, mark read, or delete a message. The direct COM SaveAs API does not provide a cross-process atomic no-clobber transaction, so concurrent writers to the same new path must coordinate externally.
- All examples must run as modules from the repository root and contain no personal paths, subjects, folders, or addresses.

## Review Focus

- Non-UTF-8 MSG HTML bytes: pass the bytes directly to Beautiful Soup and preserve characters identified by an HTML charset declaration.
- Nested layout/data tables: assign every table an absolute source index while excluding child-table text and rows from the parent.
- Combined rowspan, colspan, malformed overlap, and ragged rows: preserve logical columns internally, then return a rectangular public Pydantic matrix with empty strings distinct from structural None padding.
- Duplicate user-facing EmailKey values: reject the live batch before parsing instead of silently overwriting a dictionary entry.
- Save collisions and COM failures: never overwrite an existing MSG and always chain the underlying COM error into MessageSaveError.

## File Structure

- Create schemas/table.py for TableSelector and ExtractedTable only.
- Create extractors/_html_grid.py for provider-independent DOM traversal, text cleanup, nested-table isolation, and span expansion.
- Create extractors/html_tables.py for header inference, selector semantics, and conversion from internal grids to Pydantic results.
- Create extractors/msg_tables.py for path validation, extract-msg lifecycle management, and delegation to the shared parser.
- Create protocols/table_provider.py and protocols/message_saver.py as additive capabilities.
- Modify providers/outlook.py only for the two Outlook-specific public methods and olMSGUnicode constant.
- Add focused tests under tests/schemas, tests/extractors, tests/providers, and tests/examples.
- Add a synthetic tests/fixtures/table_message.msg plus its deterministic, unsent Outlook generator.
- Add an importable examples package with one script per requested workflow and shared display formatting.
- Update README.md, pyproject.toml, uv.lock, package exports, and domain exceptions.

---

### Task 1: Dependencies, table models, exceptions, and stable exports

**Files:**
- Modify: pyproject.toml:7-10
- Modify: uv.lock
- Create: schemas/table.py
- Modify: schemas/__init__.py:1-14
- Modify: exceptions.py:1-25
- Create: tests/schemas/test_table.py
- Create: tests/test_existing_models_unchanged.py

**Interfaces:**
- Consumes: Pydantic BaseModel, Field, field_validator, model_validator.
- Produces: TableSelector(source_index, required_columns, occurrence, header_row), ExtractedTable(source_index, caption, columns, rows), and four domain exceptions used by every later task.

- [ ] **Step 1: Add runtime and development dependency declarations**

Replace the dependency section with:

~~~toml
dependencies = [
    "beautifulsoup4>=4.13,<5",
    "extract-msg>=0.56,<0.57",
    "pydantic>=2.0",
    "pywin32>=311; sys_platform == 'win32'",
]

[dependency-groups]
dev = [
    "pytest>=8.0,<9",
]
~~~

Run:

    uv lock
    uv sync --group dev

Expected: uv.lock includes Pydantic, Beautiful Soup, extract-msg, pytest, and a Windows marker for pywin32.

- [ ] **Step 2: Write failing model and compatibility tests**

Create tests/schemas/test_table.py with concrete cases:

~~~python
import pytest
from pydantic import ValidationError

from schemas.table import ExtractedTable, TableSelector


def test_selector_defaults_select_all_tables():
    selector = TableSelector()
    assert selector.source_index is None
    assert selector.required_columns == []
    assert selector.occurrence is None
    assert selector.header_row == "auto"


def test_selector_strips_columns_and_rejects_normalized_duplicates():
    assert TableSelector(required_columns=[" Account "]).required_columns == ["Account"]
    with pytest.raises(ValidationError, match="unique"):
        TableSelector(required_columns=["Account", " account "])


@pytest.mark.parametrize("columns", [[""], ["   "]])
def test_selector_rejects_blank_required_columns(columns):
    with pytest.raises(ValidationError, match="blank"):
        TableSelector(required_columns=columns)


def test_selector_rejects_two_position_spaces():
    with pytest.raises(ValidationError, match="mutually exclusive"):
        TableSelector(source_index=0, occurrence=0)


@pytest.mark.parametrize(
    ("field", "value"),
    [("source_index", -1), ("occurrence", -1), ("header_row", -1)],
)
def test_selector_rejects_negative_indices(field, value):
    with pytest.raises(ValidationError):
        TableSelector(**{field: value})


def test_extracted_table_requires_rectangular_rows():
    with pytest.raises(ValidationError, match="same width"):
        ExtractedTable(
            source_index=0,
            columns=["A", "B"],
            rows=[["only one"]],
        )
~~~

Create tests/test_existing_models_unchanged.py:

~~~python
from datetime import datetime

from schemas.filter import SearchQuery
from schemas.result import EmailRecord, ExtractionResult, QueryResult


def test_existing_pydantic_field_sets_are_unchanged():
    assert set(EmailRecord.model_fields) == {
        "subject", "sender", "sender_email", "received_time", "body", "attachments"
    }
    assert set(SearchQuery.model_fields) == {
        "name", "email_filters", "attachment_filter", "body_filter"
    }
    assert set(QueryResult.model_fields) == {"query", "records"}
    assert set(ExtractionResult.model_fields) == {
        "name", "emails", "attachments", "bodies", "error"
    }


def test_existing_models_keep_exact_serialized_shapes_and_defaults():
    received = datetime(2026, 9, 23, 10, 0)
    email = EmailRecord(
        sender_email="sender@example.com",
        received_time=received,
    )
    assert email.model_dump() == {
        "subject": "",
        "sender": "",
        "sender_email": "sender@example.com",
        "received_time": received,
        "body": "",
        "attachments": [],
    }

    query = SearchQuery(name="daily", email_filters=[])
    assert query.model_dump() == {
        "name": "daily",
        "email_filters": [],
        "attachment_filter": None,
        "body_filter": None,
    }

    result = ExtractionResult(
        name="daily",
        emails=[email],
        attachments={"report-key": {"report.csv": b"a,b\n1,2\n"}},
        bodies={"report-key": "plain body"},
    )
    assert result.model_dump() == {
        "name": "daily",
        "emails": [email.model_dump()],
        "attachments": {"report-key": {"report.csv": b"a,b\n1,2\n"}},
        "bodies": {"report-key": "plain body"},
        "error": None,
    }
~~~

- [ ] **Step 3: Run the tests and confirm the feature is absent**

Run:

    uv run pytest tests/schemas/test_table.py tests/test_existing_models_unchanged.py -q

Expected: collection fails because schemas.table does not exist.

- [ ] **Step 4: Implement the two table models**

Create schemas/table.py:

~~~python
"""Validated selectors and results for HTML table extraction."""

from typing import Literal, Self

from pydantic import BaseModel, Field, field_validator, model_validator


def _comparison_name(value: str) -> str:
    return " ".join(value.split()).casefold()


class TableSelector(BaseModel):
    """Select tables by absolute position, columns, or filtered occurrence."""

    source_index: int | None = Field(default=None, ge=0)
    required_columns: list[str] = Field(default_factory=list)
    occurrence: int | None = Field(default=None, ge=0)
    header_row: int | Literal["auto"] = "auto"

    @field_validator("required_columns")
    @classmethod
    def validate_required_columns(cls, values: list[str]) -> list[str]:
        stripped = [value.strip() for value in values]
        if any(not value for value in stripped):
            raise ValueError("required columns cannot contain blank names")
        normalized = [_comparison_name(value) for value in stripped]
        if len(normalized) != len(set(normalized)):
            raise ValueError("required columns must be unique after normalization")
        return stripped

    @field_validator("header_row")
    @classmethod
    def validate_header_row(cls, value: int | Literal["auto"]):
        if isinstance(value, int) and value < 0:
            raise ValueError("header_row must be non-negative")
        return value

    @model_validator(mode="after")
    def validate_position_space(self) -> Self:
        if self.source_index is not None and self.occurrence is not None:
            raise ValueError("source_index and occurrence are mutually exclusive")
        return self


class ExtractedTable(BaseModel):
    """A rectangular, source-ordered table extracted from an HTML body."""

    source_index: int = Field(ge=0)
    caption: str | None = None
    columns: list[str]
    rows: list[list[str | None]]

    @model_validator(mode="after")
    def validate_rectangular_rows(self) -> Self:
        width = len(self.columns)
        if any(len(row) != width for row in self.rows):
            raise ValueError("all rows must have the same width as columns")
        return self
~~~

Re-export both names from schemas/__init__.py without modifying existing exports.

- [ ] **Step 5: Add exact domain exceptions**

Append to exceptions.py:

~~~python
class DuplicateEmailKeysError(ValueError):
    """Raised when input records would overwrite the same result key."""


class TableExtractionError(RuntimeError):
    """Raised when HTML access or table parsing fails unexpectedly."""


class MessageFileReadError(IOError):
    """Raised when a saved MSG file cannot be decoded or parsed."""


class MessageSaveError(IOError):
    """Raised when Outlook cannot save a cached message as MSG."""
~~~

- [ ] **Step 6: Run focused tests**

Run:

    uv run pytest tests/schemas/test_table.py tests/test_existing_models_unchanged.py -q

Expected: all tests pass.

- [ ] **Step 7: Commit the models and dependency baseline**

    git add pyproject.toml uv.lock schemas/table.py schemas/__init__.py exceptions.py tests/schemas/test_table.py tests/test_existing_models_unchanged.py
    git commit -m "feat: add table extraction models"

---

### Task 2: Provider-independent HTML grid normalization

**Files:**
- Create: extractors/__init__.py
- Create: extractors/_html_grid.py
- Create: tests/extractors/test_html_grid.py

**Interfaces:**
- Consumes: html: str | bytes.
- Produces: parse_html_grids(html) -> list[HtmlTableGrid], with HtmlTableGrid(source_index, caption, rows) and variable-width internal HtmlGridRow(values, has_cells, has_header_cell, in_thead). Public rectangular padding belongs to Task 3.

- [ ] **Step 1: Write failing structural-normalization tests**

Create tests/extractors/test_html_grid.py:

~~~python
import pytest

from extractors._html_grid import parse_html_grids


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
~~~

- [ ] **Step 2: Run the grid tests and confirm the module is missing**

Run:

    uv run pytest tests/extractors/test_html_grid.py -q

Expected: collection fails because extractors._html_grid does not exist.

- [ ] **Step 3: Implement immutable internal grid types and DOM helpers**

Create extractors/_html_grid.py with:

~~~python
"""Internal HTML-to-grid normalization for table extraction."""

from dataclasses import dataclass

from bs4 import BeautifulSoup, NavigableString, Tag


@dataclass(frozen=True)
class HtmlGridRow:
    values: tuple[str | None, ...]
    has_cells: bool
    has_header_cell: bool
    in_thead: bool


@dataclass(frozen=True)
class HtmlTableGrid:
    source_index: int
    caption: str | None
    rows: tuple[HtmlGridRow, ...]


def _collapse(value: str) -> str:
    return " ".join(value.split())


def _nearest_table(node: Tag | NavigableString) -> Tag | None:
    return node.find_parent("table")


def _text_without_nested_content(node: Tag, table: Tag) -> str:
    parts: list[str] = []
    for descendant in node.descendants:
        if not isinstance(descendant, NavigableString):
            continue
        if descendant.find_parent(["script", "style"]) is not None:
            continue
        if _nearest_table(descendant) is not table:
            continue
        parts.append(str(descendant))
    return _collapse(" ".join(parts))


def _logical_rows(table: Tag) -> list[Tag]:
    return [
        row
        for row in table.find_all("tr")
        if row.find_parent("table") is table
    ]


def _direct_cells(row: Tag) -> list[Tag]:
    return list(row.find_all(["th", "td"], recursive=False))


def _span(cell: Tag, name: str) -> int:
    try:
        value = int(cell.get(name, 1))
    except (TypeError, ValueError):
        return 1
    return value if value > 0 else 1


def _next_free_run(occupied: dict[int, str], start: int, width: int) -> int:
    while any(index in occupied for index in range(start, start + width)):
        start += 1
    return start
~~~

- [ ] **Step 4: Implement span expansion without losing empty-cell identity**

Add a private _expand_rows(table, rows) function that maintains active rowspans by logical column. Use this exact state transition:

~~~python
def _expand_rows(table: Tag, rows: list[Tag]) -> list[HtmlGridRow]:
    active: dict[int, tuple[str, int]] = {}
    expanded: list[HtmlGridRow] = []

    for row in rows:
        occupied: dict[int, str] = {}
        next_active: dict[int, tuple[str, int]] = {}

        for column, (value, remaining) in active.items():
            occupied[column] = value
            if remaining > 1:
                next_active[column] = (value, remaining - 1)

        column = 0
        cells = _direct_cells(row)
        for cell in cells:
            value = _text_without_nested_content(cell, table)
            colspan = _span(cell, "colspan")
            rowspan = _span(cell, "rowspan")

            column = _next_free_run(occupied, column, colspan)
            for logical_column in range(column, column + colspan):
                occupied[logical_column] = value
                if rowspan > 1:
                    next_active[logical_column] = (value, rowspan - 1)
            column += colspan

        width = max(occupied, default=-1) + 1
        values = tuple(occupied.get(index) for index in range(width))
        thead = row.find_parent("thead")
        expanded.append(
            HtmlGridRow(
                values=values,
                has_cells=bool(cells),
                has_header_cell=any(cell.name == "th" for cell in cells),
                in_thead=thead is not None and thead.find_parent("table") is table,
            )
        )
        active = next_active

    # Rows deliberately remain variable-width here. The public facade pads only
    # the selected header and retained data rows, so a discarded wide title row
    # cannot widen the returned table.
    return expanded
~~~

Implement parse_html_grids:

~~~python
def parse_html_grids(html: str | bytes) -> list[HtmlTableGrid]:
    if not isinstance(html, (str, bytes)):
        raise TypeError("html must be str or bytes")
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    grids: list[HtmlTableGrid] = []
    for source_index, table in enumerate(soup.find_all("table")):
        rows = _logical_rows(table)
        if not any(_direct_cells(row) for row in rows):
            continue
        caption_node = table.find("caption", recursive=False)
        caption = (
            _text_without_nested_content(caption_node, table)
            if isinstance(caption_node, Tag)
            else None
        )
        grids.append(
            HtmlTableGrid(
                source_index=source_index,
                caption=caption or None,
                rows=tuple(_expand_rows(table, rows)),
            )
        )
    return grids
~~~

Create extractors/__init__.py with a package docstring only at this stage.

- [ ] **Step 5: Run structural tests**

Run:

    uv run pytest tests/extractors/test_html_grid.py -q

Expected: all tests pass, including charset handling, nesting, spans, and deliberate internal ragged widths.

- [ ] **Step 6: Commit the grid normalizer**

    git add extractors/__init__.py extractors/_html_grid.py tests/extractors/test_html_grid.py
    git commit -m "feat: normalize html tables into grids"

---

### Task 3: Public HTML table selector and result facade

**Files:**
- Create: extractors/html_tables.py
- Modify: extractors/__init__.py
- Create: tests/extractors/test_html_tables.py

**Interfaces:**
- Consumes: parse_html_grids(html: str | bytes) and TableSelector.
- Produces: extract_html_tables(html: str | bytes, selector: TableSelector | None = None) -> list[ExtractedTable].

- [ ] **Step 1: Write failing public-parser tests**

Create tests/extractors/test_html_tables.py with:

~~~python
import pytest

from exceptions import TableExtractionError
from extractors.html_tables import extract_html_tables
from schemas.table import TableSelector


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
    html = "<table><tr><td>title</td></tr><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>"
    table = extract_html_tables(html, TableSelector(header_row=1))[0]
    assert table.columns == ["A", "B"]
    assert table.rows == [["1", "2"]]


def test_explicit_header_accepts_a_structurally_present_blank_cell():
    html = "<table><tr><th></th></tr><tr><td>x</td></tr></table>"
    table = extract_html_tables(html, TableSelector(header_row=0))[0]
    assert table.columns == [""]
    assert table.rows == [["x"]]


def test_auto_header_uses_first_direct_th_row_without_thead():
    html = "<table><tr><td>title</td></tr><tr><th>A</th></tr><tr><td>1</td></tr></table>"
    table = extract_html_tables(html)[0]
    assert table.columns == ["A"]
    assert table.rows == [["1"]]


def test_auto_header_falls_back_to_first_nonempty_row_without_th():
    html = "<table><tr><td></td></tr><tr><td>A</td></tr><tr><td>1</td></tr></table>"
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
    html = '<meta charset="utf-8"><table><tr><th>Caf\u00e9</th></tr><tr><td>yes</td></tr></table>'
    assert extract_html_tables(html) == extract_html_tables(html.encode("utf-8"))


def test_unexpected_grid_failure_is_chained(monkeypatch):
    def fail(_html):
        raise RuntimeError("parser exploded")

    monkeypatch.setattr("extractors.html_tables.parse_html_grids", fail)
    with pytest.raises(TableExtractionError, match="parser exploded") as raised:
        extract_html_tables("<table></table>")
    assert isinstance(raised.value.__cause__, RuntimeError)
~~~

- [ ] **Step 2: Run the tests and confirm the facade is absent**

Run:

    uv run pytest tests/extractors/test_html_tables.py -q

Expected: collection fails because extractors.html_tables does not exist.

- [ ] **Step 3: Implement deterministic header selection and table construction**

Create extractors/html_tables.py:

~~~python
"""Public HTML table extraction and selector semantics."""

from exceptions import TableExtractionError
from schemas.table import ExtractedTable, TableSelector

from ._html_grid import HtmlGridRow, HtmlTableGrid, parse_html_grids


def _has_text(row: HtmlGridRow) -> bool:
    return any(value not in (None, "") for value in row.values)


def _header_index(grid: HtmlTableGrid, selector: TableSelector) -> int | None:
    if isinstance(selector.header_row, int):
        if selector.header_row >= len(grid.rows):
            return None
        return selector.header_row if grid.rows[selector.header_row].has_cells else None

    thead_rows = [
        index
        for index, row in enumerate(grid.rows)
        if row.in_thead and _has_text(row)
    ]
    if thead_rows:
        return thead_rows[-1]

    for index, row in enumerate(grid.rows):
        if row.has_header_cell:
            return index
    for index, row in enumerate(grid.rows):
        if _has_text(row):
            return index
    return None


def _to_table(grid: HtmlTableGrid, selector: TableSelector) -> ExtractedTable | None:
    header_index = _header_index(grid, selector)
    if header_index is None:
        return None
    header = grid.rows[header_index]
    data = [
        row
        for row in grid.rows[header_index + 1 :]
        if _has_text(row)
    ]
    width = max(
        [len(header.values), *(len(row.values) for row in data)],
        default=0,
    )
    columns = [
        value if value is not None else ""
        for value in header.values
    ] + [""] * (width - len(header.values))
    rows = [
        list(row.values) + [None] * (width - len(row.values))
        for row in data
    ]
    return ExtractedTable(
        source_index=grid.source_index,
        caption=grid.caption,
        columns=columns,
        rows=rows,
    )


def _comparison_name(value: str) -> str:
    return " ".join(value.split()).casefold()


def _has_required_columns(table: ExtractedTable, required: list[str]) -> bool:
    available = {_comparison_name(column) for column in table.columns}
    return all(_comparison_name(column) in available for column in required)
~~~

- [ ] **Step 4: Implement source-index and filtered-occurrence ordering**

Add:

~~~python
def extract_html_tables(
    html: str | bytes,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    """Extract structured tables from an HTML string or byte sequence."""
    if not isinstance(html, (str, bytes)):
        raise TypeError("html must be str or bytes")
    selected = selector or TableSelector()
    try:
        grids = parse_html_grids(html)
        if selected.source_index is not None:
            grids = [
                grid for grid in grids
                if grid.source_index == selected.source_index
            ]

        tables = [
            table
            for grid in grids
            if (table := _to_table(grid, selected)) is not None
            and _has_required_columns(table, selected.required_columns)
        ]
    except TableExtractionError:
        raise
    except Exception as exc:
        raise TableExtractionError(f"Failed to extract HTML tables: {exc}") from exc

    if selected.occurrence is None:
        return tables
    if selected.occurrence >= len(tables):
        return []
    return [tables[selected.occurrence]]
~~~

Ensure TypeError from invalid html input remains TypeError. Export extract_html_tables from extractors/__init__.py.

- [ ] **Step 5: Add review-focus regression cases**

Extend tests/extractors/test_html_tables.py with a nested table whose parent has no usable text, combined rowspan/colspan data rows, duplicate/blank headers, and an out-of-range occurrence:

~~~python
def test_out_of_range_occurrence_is_empty():
    assert extract_html_tables(HTML, TableSelector(occurrence=99)) == []


def test_duplicate_and_blank_header_labels_are_preserved():
    html = "<table><tr><th>A</th><th></th><th>A</th></tr><tr><td>1</td><td></td><td>3</td></tr></table>"
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
~~~

- [ ] **Step 6: Run all pure-parser tests**

Run:

    uv run pytest tests/schemas/test_table.py tests/extractors/test_html_grid.py tests/extractors/test_html_tables.py -q

Expected: all tests pass.

- [ ] **Step 7: Commit the public parser**

    git add extractors/__init__.py extractors/html_tables.py tests/extractors/test_html_tables.py
    git commit -m "feat: select structured tables from html"

---

### Task 4: Outlook-independent MSG reader and real offline fixture

**Files:**
- Create: extractors/msg_tables.py
- Modify: extractors/__init__.py
- Create: tests/extractors/test_msg_tables.py
- Create: tests/fixtures/generate_table_message.py
- Create: tests/fixtures/table_message.msg

**Interfaces:**
- Consumes: extract_msg.openMsg(path), message.htmlBody: bytes | None, extract_html_tables().
- Produces: extract_tables_from_msg(path: str | Path, selector: TableSelector | None = None) -> list[ExtractedTable].

- [ ] **Step 1: Write failing path, lifecycle, and delegation tests**

Create tests/extractors/test_msg_tables.py:

~~~python
from pathlib import Path
import subprocess
import sys

import pytest

from exceptions import MessageFileReadError
from extractors.msg_tables import extract_tables_from_msg
from schemas.table import TableSelector


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
    message = FakeMessage(b"<table><tr><th>A</th></tr><tr><td>1</td></tr></table>")
    monkeypatch.setattr("extractors.msg_tables.extract_msg.openMsg", lambda _path: message)

    tables = extract_tables_from_msg(path, TableSelector(required_columns=["A"]))

    assert tables[0].rows == [["1"]]
    assert message.closed is True


def test_delegates_the_exact_msg_html_bytes_and_selector(tmp_path, monkeypatch):
    path = tmp_path / "raw.msg"
    path.write_bytes(b"placeholder")
    html_body = b"<meta charset=windows-1252><table><tr><td>Caf\xe9</td></tr></table>"
    message = FakeMessage(html_body)
    selector = TableSelector(source_index=0)
    monkeypatch.setattr(
        "extractors.msg_tables.extract_msg.openMsg",
        lambda _path: message,
    )

    def capture(html, received_selector):
        assert html is html_body
        assert received_selector is selector
        return []

    monkeypatch.setattr("extractors.msg_tables.extract_html_tables", capture)
    assert extract_tables_from_msg(path, selector) == []
    assert message.closed is True


def test_empty_msg_html_is_successful_empty_result(tmp_path, monkeypatch):
    path = tmp_path / "empty.msg"
    path.write_bytes(b"placeholder")
    monkeypatch.setattr(
        "extractors.msg_tables.extract_msg.openMsg",
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
    with pytest.raises(ValueError, match=r"\\.msg"):
        extract_tables_from_msg(wrong_suffix)


def test_unexpected_path_access_failure_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "permission.msg"

    def fail(_path):
        raise PermissionError("access denied")

    monkeypatch.setattr("extractors.msg_tables._validated_msg_path", fail)
    with pytest.raises(MessageFileReadError, match="permission.msg") as raised:
        extract_tables_from_msg(path)
    assert isinstance(raised.value.__cause__, PermissionError)


def test_corrupt_msg_error_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "bad.msg"
    path.write_bytes(b"bad")

    def fail(_path):
        raise RuntimeError("bad compound file")

    monkeypatch.setattr("extractors.msg_tables.extract_msg.openMsg", fail)
    with pytest.raises(MessageFileReadError, match="bad.msg") as raised:
        extract_tables_from_msg(path)
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_html_body_failure_closes_message_and_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "broken-body.msg"
    path.write_bytes(b"placeholder")
    message = FakeMessage(body_failure=RuntimeError("body unavailable"))
    monkeypatch.setattr("extractors.msg_tables.extract_msg.openMsg", lambda _path: message)

    with pytest.raises(MessageFileReadError, match="broken-body.msg") as raised:
        extract_tables_from_msg(path)

    assert message.closed is True
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_parser_failure_closes_message_and_is_chained(tmp_path, monkeypatch):
    path = tmp_path / "parser-failure.msg"
    path.write_bytes(b"placeholder")
    message = FakeMessage(b"<table></table>")
    monkeypatch.setattr("extractors.msg_tables.extract_msg.openMsg", lambda _path: message)

    def fail(_html, _selector):
        raise RuntimeError("parser exploded")

    monkeypatch.setattr("extractors.msg_tables.extract_html_tables", fail)
    with pytest.raises(MessageFileReadError, match="parser-failure.msg") as raised:
        extract_tables_from_msg(path)

    assert message.closed is True
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_msg_reader_import_does_not_import_outlook_dependencies():
    code = (
        "import sys; import extractors.msg_tables; "
        "assert 'win32com' not in sys.modules; "
        "assert 'providers.outlook' not in sys.modules"
    )
    subprocess.run([sys.executable, "-c", code], check=True)
~~~

- [ ] **Step 2: Run focused tests and confirm the adapter is absent**

Run:

    uv run pytest tests/extractors/test_msg_tables.py -q

Expected: collection fails because extractors.msg_tables does not exist.

- [ ] **Step 3: Implement validated, context-managed MSG reading**

Create extractors/msg_tables.py:

~~~python
"""Extract HTML tables from saved Outlook MSG files without Outlook."""

from pathlib import Path

import extract_msg

from exceptions import MessageFileReadError
from schemas.table import ExtractedTable, TableSelector

from .html_tables import extract_html_tables


def _validated_msg_path(path: str | Path) -> Path:
    candidate = Path(path)
    if not candidate.exists():
        raise FileNotFoundError(candidate)
    if not candidate.is_file():
        raise IsADirectoryError(candidate)
    if candidate.suffix.casefold() != ".msg":
        raise ValueError(f"Expected a .msg file, got '{candidate}'")
    return candidate


def extract_tables_from_msg(
    path: str | Path,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    """Extract selected tables from one saved Outlook MSG file."""
    candidate = Path(path)
    try:
        msg_path = _validated_msg_path(candidate)
    except (FileNotFoundError, IsADirectoryError, ValueError):
        raise
    except OSError as exc:
        raise MessageFileReadError(
            f"Failed to access MSG file '{candidate}': {exc}"
        ) from exc

    try:
        with extract_msg.openMsg(str(msg_path)) as message:
            html_body = message.htmlBody
            if not html_body:
                return []
            return extract_html_tables(html_body, selector)
    except Exception as exc:
        raise MessageFileReadError(
            f"Failed to read tables from MSG file '{msg_path}': {exc}"
        ) from exc
~~~

Keep the three ordinary validation outcomes unchanged, but chain unexpected path-access
errors and all open/body/parser failures as MessageFileReadError. Export
extract_tables_from_msg from extractors/__init__.py.

- [ ] **Step 4: Generate a sanitized table fixture without sending mail**

Create tests/fixtures/generate_table_message.py:

~~~python
"""Generate the synthetic MSG fixture; requires desktop Outlook once."""

from pathlib import Path

import win32com.client


HTML = """
<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252"></head><body>
<table>
  <tr><th>Item</th><th>Quantity</th></tr>
  <tr><td>Caf\u00e9</td><td>2</td></tr>
</table>
<table>
  <tr><th>Account</th><th>Amount</th></tr>
  <tr><td>A-100</td><td>125.50</td></tr>
</table>
</body></html>
"""


def main() -> None:
    destination = Path(__file__).with_name("table_message.msg").resolve()
    if destination.exists():
        raise FileExistsError(f"Refusing to overwrite fixture: {destination}")
    outlook = win32com.client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Subject = "Synthetic table fixture"
    message.HTMLBody = HTML
    message.SaveAs(str(destination), 9)
    print(destination)


if __name__ == "__main__":
    main()
~~~

Run once on the Windows development host:

    uv run python tests/fixtures/generate_table_message.py

Expected: tests/fixtures/table_message.msg exists; no message is sent or added to a mailbox.

- [ ] **Step 5: Add the true offline fixture test**

Append:

~~~python
def test_real_msg_fixture_extracts_two_tables_without_outlook():
    fixture = Path(__file__).parents[1] / "fixtures" / "table_message.msg"
    tables = extract_tables_from_msg(fixture)
    assert [table.columns for table in tables] == [
        ["Item", "Quantity"],
        ["Account", "Amount"],
    ]
    assert tables[0].rows == [["Caf\u00e9", "2"]]
    assert tables[1].rows == [["A-100", "125.50"]]
~~~

- [ ] **Step 6: Run MSG tests in a fresh process**

Run:

    uv run pytest tests/extractors/test_msg_tables.py -q

Expected: all tests pass and the subprocess assertion proves neither win32com nor providers.outlook is imported.

- [ ] **Step 7: Commit the standalone reader and sanitized fixture**

    git add extractors/__init__.py extractors/msg_tables.py tests/extractors/test_msg_tables.py tests/fixtures/generate_table_message.py tests/fixtures/table_message.msg
    git commit -m "feat: extract tables from msg files"

---

### Task 5: Live Outlook table extraction capability

**Files:**
- Create: protocols/table_provider.py
- Modify: protocols/__init__.py:1-5
- Modify: providers/outlook.py:1-130
- Create: tests/providers/test_outlook_tables.py
- Create: tests/providers/test_existing_outlook_behavior.py

**Interfaces:**
- Consumes: EmailRecord cache keys, MailItem.HTMLBody, TableSelector, extract_html_tables().
- Produces: OutlookProvider.extract_tables(email_records, selector) -> dict[EmailKey, list[ExtractedTable]] and TableExtractingEmailProvider.

- [ ] **Step 1: Write failing live-adapter tests with fake COM objects**

Create tests/providers/test_outlook_tables.py:

~~~python
from datetime import datetime

import pytest

from exceptions import (
    DuplicateEmailKeysError,
    EmailNotInCacheError,
    TableExtractionError,
)
from providers.outlook import OutlookProvider
from schemas.result import EmailRecord
from schemas.table import TableSelector


def record(sender="sender@example.com", subject="Report", minute=0):
    return EmailRecord(
        subject=subject,
        sender="Sender",
        sender_email=sender,
        received_time=datetime(2026, 9, 23, 10, minute),
        body="plain text remains unchanged",
    )


def provider_with_messages(records_and_messages):
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {
        provider._get_cache_key(item): message
        for item, message in records_and_messages
    }
    return provider


def test_extracts_live_html_and_preserves_plain_body():
    item = record()
    message = type(
        "Message",
        (),
        {"HTMLBody": "<table><tr><th>A</th></tr><tr><td>1</td></tr></table>"},
    )()
    provider = provider_with_messages([(item, message)])

    result = provider.extract_tables([item], TableSelector(required_columns=["A"]))

    assert result[provider._make_email_key(item)][0].rows == [["1"]]
    assert item.body == "plain text remains unchanged"


@pytest.mark.parametrize("html_body", [None, ""])
def test_live_empty_html_keeps_email_key(html_body):
    item = record()
    provider = provider_with_messages([
        (item, type("Message", (), {"HTMLBody": html_body})()),
    ])
    assert provider.extract_tables([item]) == {provider._make_email_key(item): []}


@pytest.mark.parametrize("html_body", [b"<table></table>", 123])
def test_live_html_body_requires_com_string(html_body):
    item = record()
    message = type("Message", (), {"HTMLBody": html_body})()
    provider = provider_with_messages([(item, message)])

    with pytest.raises(TableExtractionError, match="Report") as raised:
        provider.extract_tables([item])

    assert isinstance(raised.value.__cause__, TypeError)


def test_missing_cached_message_is_explicit():
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    with pytest.raises(EmailNotInCacheError):
        provider.extract_tables([record()])


def test_duplicate_user_facing_keys_fail_before_overwrite():
    first = record("first@example.com")
    second = record("second@example.com")
    provider = provider_with_messages([
        (first, type("Message", (), {"HTMLBody": ""})()),
        (second, type("Message", (), {"HTMLBody": ""})()),
    ])
    with pytest.raises(DuplicateEmailKeysError):
        provider.extract_tables([first, second])


def test_html_property_failure_is_chained():
    class BrokenMessage:
        @property
        def HTMLBody(self):
            raise RuntimeError("COM failure")

    item = record()
    provider = provider_with_messages([(item, BrokenMessage())])
    with pytest.raises(TableExtractionError, match="Report") as raised:
        provider.extract_tables([item])
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_shared_parser_failure_gets_email_context_and_is_chained(monkeypatch):
    item = record()
    message = type("Message", (), {"HTMLBody": "<table></table>"})()
    provider = provider_with_messages([(item, message)])
    parser_error = TableExtractionError("parser exploded")

    def fail(_html, _selector):
        raise parser_error

    monkeypatch.setattr("providers.outlook.extract_html_tables", fail)
    with pytest.raises(TableExtractionError, match="Report") as raised:
        provider.extract_tables([item])

    assert raised.value.__cause__ is parser_error


def test_preserves_email_order_and_table_source_order():
    first = record(subject="First")
    second = record(subject="Second")
    first_message = type(
        "Message",
        (),
        {"HTMLBody": "<table><tr><th>A</th></tr></table><table><tr><th>B</th></tr></table>"},
    )()
    second_message = type(
        "Message",
        (),
        {"HTMLBody": "<table><tr><th>C</th></tr></table>"},
    )()
    provider = provider_with_messages([
        (first, first_message),
        (second, second_message),
    ])

    result = provider.extract_tables([first, second])

    assert list(result) == [
        provider._make_email_key(first),
        provider._make_email_key(second),
    ]
    assert [table.source_index for table in result[provider._make_email_key(first)]] == [0, 1]


def test_first_failure_does_not_touch_later_message():
    class BrokenFirst:
        @property
        def HTMLBody(self):
            raise RuntimeError("first failed")

    class LaterTripwire:
        touched = False

        @property
        def HTMLBody(self):
            self.touched = True
            raise AssertionError("later message must not be touched")

    first = record(subject="First")
    second = record(subject="Second")
    later = LaterTripwire()
    provider = provider_with_messages([(first, BrokenFirst()), (second, later)])

    with pytest.raises(TableExtractionError, match="First"):
        provider.extract_tables([first, second])

    assert later.touched is False
~~~

Before changing OutlookProvider, also capture the current non-table behavior in
tests/providers/test_existing_outlook_behavior.py so the additive methods cannot
silently alter filtering, plain-body matching, or attachment reads:

~~~python
from datetime import datetime
from pathlib import Path

from providers.outlook import OutlookProvider
from schemas.filter import (
    AttachmentFilter,
    BodyFilter,
    FolderFilter,
    KeywordFilter,
    SearchQuery,
)
from schemas.result import EmailRecord


def record(subject, body="", attachments=None):
    return EmailRecord(
        subject=subject,
        sender="Sender",
        sender_email="sender@example.com",
        received_time=datetime(2026, 9, 23, 10, 0),
        body=body,
        attachments=attachments or [],
    )


def test_filter_emails_keeps_existing_subject_filtering(monkeypatch):
    provider = object.__new__(OutlookProvider)
    quarterly = record("Quarterly Report")
    unrelated = record("Team Lunch")
    monkeypatch.setattr(
        provider,
        "_search_folder",
        lambda *_args: [quarterly, unrelated],
    )
    query = SearchQuery(
        name="reports",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["quarterly"], exact_match=False),
        ],
    )

    assert provider.filter_emails([query])["reports"].records == [quarterly]


def test_filter_body_keeps_plain_text_body_matching():
    provider = object.__new__(OutlookProvider)
    matching = record("First", body="Contains needle")
    other = record("Second", body="No match")

    result = provider.filter_body(
        [matching, other],
        BodyFilter(keywords=["needle"], exact_match=False),
    )

    assert result == {provider._make_email_key(matching): "Contains needle"}


class FakeAttachment:
    def __init__(self, filename, content):
        self.FileName = filename
        self.content = content

    def SaveAsFile(self, path):
        Path(path).write_bytes(self.content)


class FakeAttachments:
    def __init__(self, *attachments):
        self._attachments = attachments
        self.Count = len(attachments)

    def Item(self, one_based_index):
        return self._attachments[one_based_index - 1]


def test_filter_attachments_keeps_filename_filtering_and_binary_reads(tmp_path):
    item = record(
        "Report",
        attachments=["Monthly Report.CSV", "ignore.txt"],
    )
    message = type(
        "Message",
        (),
        {
            "Attachments": FakeAttachments(
                FakeAttachment("Monthly Report.CSV", b"csv-content"),
                FakeAttachment("ignore.txt", b"ignore"),
            )
        },
    )()
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {provider._get_cache_key(item): message}
    provider._temp_dir = str(tmp_path)
    attachment_filter = AttachmentFilter(
        filters=[KeywordFilter(keywords=["report"], exact_match=False)]
    )

    assert provider.filter_attachments([item], attachment_filter) == {
        "Monthly Report.CSV": b"csv-content"
    }
~~~

- [ ] **Step 2: Run the tests and confirm the method is absent**

Run:

    uv run pytest tests/providers/test_outlook_tables.py tests/providers/test_existing_outlook_behavior.py -q

Expected: the existing-behavior tests pass, while failures in test_outlook_tables.py show OutlookProvider has no extract_tables method.

- [ ] **Step 3: Add the additive capability protocol**

Create protocols/table_provider.py:

~~~python
"""Optional protocol for providers that expose structured table extraction."""

from typing import Protocol

from schemas.result import EmailKey, EmailRecord
from schemas.table import ExtractedTable, TableSelector


class TableExtractingEmailProvider(Protocol):
    def extract_tables(
        self,
        email_records: list[EmailRecord],
        selector: TableSelector | None = None,
    ) -> dict[EmailKey, list[ExtractedTable]]:
        """Extract selected tables from the supplied email records."""
~~~

Re-export it from protocols/__init__.py. Do not edit protocols/provider.py.

- [ ] **Step 4: Implement atomic live extraction**

Import the new models, exceptions, and extract_html_tables in providers/outlook.py. The
adapter deliberately wraps even an existing TableExtractionError once so the outer error
identifies the failing email; the original parser error remains available through
`__cause__`. Add:

~~~python
def extract_tables(
    self,
    email_records: list[EmailRecord],
    selector: TableSelector | None = None,
) -> dict[EmailKey, list[ExtractedTable]]:
    """Extract selected tables from cached live Outlook messages."""
    email_keys = [self._make_email_key(record) for record in email_records]
    duplicates = {
        key for key in email_keys if email_keys.count(key) > 1
    }
    if duplicates:
        raise DuplicateEmailKeysError(
            f"Duplicate email keys found: {sorted(duplicates)}"
        )

    result: dict[EmailKey, list[ExtractedTable]] = {}
    for record, email_key in zip(email_records, email_keys, strict=True):
        message = self._message_cache.get(self._get_cache_key(record))
        if message is None:
            raise EmailNotInCacheError(
                f"Email record for '{record.subject}' not found in cache. "
                "Make sure to call filter_emails() first."
            )
        try:
            html_body = message.HTMLBody
            if html_body is None or html_body == "":
                result[email_key] = []
                continue
            if not isinstance(html_body, str):
                raise TypeError(
                    f"Outlook HTMLBody must be str, got {type(html_body).__name__}"
                )
            result[email_key] = extract_html_tables(
                html_body,
                selector,
            )
        except Exception as exc:
            raise TableExtractionError(
                f"Failed to extract tables from '{record.subject}': {exc}"
            ) from exc
    return result
~~~

- [ ] **Step 5: Run live-adapter and existing-model tests**

Run:

    uv run pytest tests/providers/test_outlook_tables.py tests/providers/test_existing_outlook_behavior.py tests/test_existing_models_unchanged.py -q

Expected: all tests pass; existing filtering, plain-body matching, attachment reads, and model fields remain unchanged.

- [ ] **Step 6: Commit live table extraction**

    git add protocols/table_provider.py protocols/__init__.py providers/outlook.py tests/providers/test_outlook_tables.py tests/providers/test_existing_outlook_behavior.py
    git commit -m "feat: extract tables from live outlook mail"

---

### Task 6: Safe Unicode MSG save utility

**Files:**
- Create: protocols/message_saver.py
- Modify: protocols/__init__.py
- Modify: providers/outlook.py:16-18 and 154-259
- Create: tests/providers/test_outlook_message_save.py

**Interfaces:**
- Consumes: one cached EmailRecord and destination str | Path.
- Produces: OutlookProvider.save_message(email_record, save_path) -> Path and MessageSavingEmailProvider.

- [ ] **Step 1: Write failing save-contract tests**

Create tests/providers/test_outlook_message_save.py:

~~~python
from datetime import datetime
from pathlib import Path

import pytest

from exceptions import (
    EmailNotInCacheError,
    InvalidSavePathError,
    MessageSaveError,
)
from providers.outlook import OutlookProvider
from schemas.result import EmailRecord


class FakeMessage:
    def __init__(self, failure=None):
        object.__setattr__(self, "calls", [])
        object.__setattr__(self, "failure", failure)

    def __setattr__(self, name, value):
        raise AssertionError(f"save_message must not assign message.{name}")

    def SaveAs(self, path, save_type):
        self.calls.append((path, save_type))
        if self.failure:
            raise self.failure

    def __getattr__(self, name):
        if name in {"Send", "Move", "Delete", "Save"}:
            raise AssertionError(f"save_message must not call {name}")
        raise AttributeError(name)


def record():
    return EmailRecord(
        subject="Report",
        sender="Sender",
        sender_email="sender@example.com",
        received_time=datetime(2026, 9, 23, 10, 0),
    )


def cached_provider(item, message):
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {provider._get_cache_key(item): message}
    return provider


def test_saves_with_unicode_msg_type_without_mutating_message(tmp_path):
    item = record()
    message = FakeMessage()
    provider = cached_provider(item, message)
    destination = tmp_path / "report.msg"

    assert provider.save_message(item, destination) == destination
    assert message.calls == [(str(destination), 9)]


def test_refuses_existing_destination_without_calling_com(tmp_path):
    item = record()
    message = FakeMessage()
    provider = cached_provider(item, message)
    destination = tmp_path / "report.msg"
    destination.write_bytes(b"keep")

    with pytest.raises(FileExistsError):
        provider.save_message(item, destination)
    assert destination.read_bytes() == b"keep"
    assert message.calls == []


def test_existing_directory_with_msg_suffix_is_also_a_collision(tmp_path):
    item = record()
    message = FakeMessage()
    provider = cached_provider(item, message)
    destination = tmp_path / "report.msg"
    destination.mkdir()

    with pytest.raises(FileExistsError):
        provider.save_message(item, destination)
    assert message.calls == []


@pytest.mark.parametrize("name", ["report.eml", "report"])
def test_requires_msg_suffix(tmp_path, name):
    item = record()
    provider = cached_provider(item, FakeMessage())
    with pytest.raises(InvalidSavePathError, match=r"\\.msg"):
        provider.save_message(item, tmp_path / name)


def test_requires_existing_parent(tmp_path):
    item = record()
    provider = cached_provider(item, FakeMessage())
    with pytest.raises(InvalidSavePathError):
        provider.save_message(item, tmp_path / "missing" / "report.msg")


def test_requires_cached_record(tmp_path):
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    with pytest.raises(EmailNotInCacheError):
        provider.save_message(record(), tmp_path / "report.msg")


def test_com_save_failure_is_chained(tmp_path):
    item = record()
    provider = cached_provider(item, FakeMessage(RuntimeError("disk failure")))
    with pytest.raises(MessageSaveError, match="Report") as raised:
        provider.save_message(item, tmp_path / "report.msg")
    assert isinstance(raised.value.__cause__, RuntimeError)
~~~

- [ ] **Step 2: Run tests and confirm save_message is absent**

Run:

    uv run pytest tests/providers/test_outlook_message_save.py -q

Expected: failures show OutlookProvider has no save_message method.

- [ ] **Step 3: Add the separate save capability protocol**

Create protocols/message_saver.py:

~~~python
"""Optional protocol for providers that persist complete messages."""

from pathlib import Path
from typing import Protocol

from schemas.result import EmailRecord


class MessageSavingEmailProvider(Protocol):
    def save_message(
        self,
        email_record: EmailRecord,
        save_path: str | Path,
    ) -> Path:
        """Save one complete message and return its destination."""
~~~

Re-export the protocol from protocols/__init__.py without editing EmailProvider.

- [ ] **Step 4: Implement safe Outlook SaveAs**

The overwrite guard is deliberately a preflight guarantee: fail before COM when the
destination already exists. Keep the required direct SaveAs(destination, 9) call; do
not imply an atomic cross-process reservation that Outlook's API does not offer.

Add beside OL_MAIL_CLASS:

~~~python
OL_MSG_UNICODE = 9  # olMSGUnicode - Outlook Unicode message format
~~~

Add to OutlookProvider:

~~~python
def save_message(
    self,
    email_record: EmailRecord,
    save_path: str | Path,
) -> Path:
    """Save one cached live Outlook message as a Unicode MSG file."""
    destination = Path(save_path)
    if destination.suffix.casefold() != ".msg":
        raise InvalidSavePathError(
            f"Save path '{destination}' must have a .msg suffix."
        )
    if not destination.parent.exists() or not destination.parent.is_dir():
        raise InvalidSavePathError(
            f"Save directory '{destination.parent}' does not exist or is not a directory."
        )
    if destination.exists():
        raise FileExistsError(destination)

    message = self._message_cache.get(self._get_cache_key(email_record))
    if message is None:
        raise EmailNotInCacheError(
            f"Email record for '{email_record.subject}' not found in cache. "
            "Make sure to call filter_emails() first."
        )

    try:
        message.SaveAs(str(destination), OL_MSG_UNICODE)
    except Exception as exc:
        raise MessageSaveError(
            f"Failed to save '{email_record.subject}' to '{destination}': {exc}"
        ) from exc
    return destination
~~~

- [ ] **Step 5: Run save and live extraction tests**

Run:

    uv run pytest tests/providers/test_outlook_message_save.py tests/providers/test_outlook_tables.py -q

Expected: all tests pass and no fake message receives unrelated method calls.

- [ ] **Step 6: Commit the save utility**

    git add protocols/message_saver.py protocols/__init__.py providers/outlook.py tests/providers/test_outlook_message_save.py
    git commit -m "feat: save outlook messages as unicode msg"

---

### Task 7: Runnable examples and usage documentation

**Files:**
- Create: examples/__init__.py
- Create: examples/_display.py
- Create: examples/extract_tables_from_msg.py
- Create: examples/extract_live_tables.py
- Create: examples/save_and_extract_message.py
- Create: examples/README.md
- Modify: README.md
- Create: tests/examples/test_examples.py

**Interfaces:**
- Consumes: the three approved public workflows.
- Produces: runnable module examples and user-facing documentation with no pandas or private data.

- [ ] **Step 1: Write failing example import and offline execution tests**

Create tests/examples/test_examples.py:

~~~python
from pathlib import Path
import subprocess
import sys


def test_example_modules_import_without_running_cli():
    __import__("examples.extract_tables_from_msg")
    __import__("examples.extract_live_tables")
    __import__("examples.save_and_extract_message")


def test_offline_example_runs_against_fixture_without_outlook():
    fixture = Path(__file__).parents[1] / "fixtures" / "table_message.msg"
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "examples.extract_tables_from_msg",
            str(fixture),
            "--column",
            "Account",
            "--column",
            "Amount",
            "--occurrence",
            "0",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    assert "Account | Amount" in result.stdout
    assert "A-100 | 125.50" in result.stdout
~~~

- [ ] **Step 2: Run tests and confirm examples are absent**

Run:

    uv run pytest tests/examples/test_examples.py -q

Expected: imports fail because examples is not present.

- [ ] **Step 3: Add shared dependency-free display formatting**

Create examples/__init__.py with a package docstring and examples/_display.py:

~~~python
from schemas.table import ExtractedTable


def print_tables(tables: list[ExtractedTable]) -> None:
    if not tables:
        print("No matching tables found.")
        return
    for table in tables:
        print(f"Table source_index={table.source_index}")
        if table.caption:
            print(f"Caption: {table.caption}")
        print(" | ".join(table.columns))
        for row in table.rows:
            print(" | ".join("" if value is None else value for value in row))
        print()
~~~

- [ ] **Step 4: Add the standalone MSG example CLI**

Create examples/extract_tables_from_msg.py with argparse arguments:

~~~python
from argparse import ArgumentParser
from pathlib import Path

from extractors import extract_tables_from_msg
from schemas import TableSelector

from ._display import print_tables


def main() -> None:
    parser = ArgumentParser(description="Extract tables from a saved MSG without Outlook.")
    parser.add_argument("msg_path", type=Path)
    parser.add_argument("--column", action="append", default=[])
    position = parser.add_mutually_exclusive_group()
    position.add_argument("--source-index", type=int)
    position.add_argument("--occurrence", type=int)
    args = parser.parse_args()

    selector = TableSelector(
        source_index=args.source_index,
        required_columns=args.column,
        occurrence=args.occurrence,
    )
    print_tables(extract_tables_from_msg(args.msg_path, selector))


if __name__ == "__main__":
    main()
~~~

- [ ] **Step 5: Add live extraction and save-then-extract CLIs**

Create examples/extract_live_tables.py:

~~~python
from argparse import ArgumentParser

from schemas import FolderFilter, KeywordFilter, SearchQuery, TableSelector

from ._display import print_tables


def main() -> None:
    parser = ArgumentParser(description="Extract tables from live Outlook mail.")
    parser.add_argument("subject")
    parser.add_argument("--folder", default="Inbox")
    parser.add_argument("--column", action="append", default=[])
    position = parser.add_mutually_exclusive_group()
    position.add_argument("--source-index", type=int)
    position.add_argument("--occurrence", type=int)
    args = parser.parse_args()

    # Importing inside main keeps a simple module import from launching Outlook.
    from providers import OutlookProvider

    provider = OutlookProvider()
    query = SearchQuery(
        name="live-table-example",
        email_filters=[
            FolderFilter(folder_name=args.folder),
            KeywordFilter(keywords=[args.subject], exact_match=True),
        ],
    )
    records = provider.filter_emails([query])[query.name].records
    tables_by_email = provider.extract_tables(
        records,
        TableSelector(
            source_index=args.source_index,
            required_columns=args.column,
            occurrence=args.occurrence,
        ),
    )
    for email_key, tables in tables_by_email.items():
        print(email_key)
        print_tables(tables)


if __name__ == "__main__":
    main()
~~~

Create examples/save_and_extract_message.py:

~~~python
from argparse import ArgumentParser
from pathlib import Path

from extractors import extract_tables_from_msg
from schemas import FolderFilter, KeywordFilter, SearchQuery, TableSelector

from ._display import print_tables


def main() -> None:
    parser = ArgumentParser(
        description="Save one live Outlook message and replay its tables offline."
    )
    parser.add_argument("subject")
    parser.add_argument("destination", type=Path)
    parser.add_argument("--folder", default="Inbox")
    parser.add_argument("--column", action="append", default=[])
    position = parser.add_mutually_exclusive_group()
    position.add_argument("--source-index", type=int)
    position.add_argument("--occurrence", type=int)
    args = parser.parse_args()

    from providers import OutlookProvider

    provider = OutlookProvider()
    query = SearchQuery(
        name="save-table-example",
        email_filters=[
            FolderFilter(folder_name=args.folder),
            KeywordFilter(keywords=[args.subject], exact_match=True),
        ],
    )
    records = provider.filter_emails([query])[query.name].records
    if not records:
        raise SystemExit(
            f"No email with subject '{args.subject}' found in '{args.folder}'."
        )

    saved_path = provider.save_message(records[0], args.destination)
    print(f"Saved: {saved_path}")
    tables = extract_tables_from_msg(
        saved_path,
        TableSelector(
            source_index=args.source_index,
            required_columns=args.column,
            occurrence=args.occurrence,
        ),
    )
    print_tables(tables)


if __name__ == "__main__":
    main()
~~~

Both files put COM construction inside main, so importing their modules does not launch Outlook.

- [ ] **Step 6: Add exact commands and compatibility explanation**

Create examples/README.md with these commands:

    uv run python -m examples.extract_tables_from_msg path/to/report.msg --column Account --column Amount --occurrence 0
    uv run python -m examples.extract_live_tables "Daily Report" --folder Inbox --column Account --column Amount --occurrence 0
    uv run python -m examples.save_and_extract_message "Daily Report" path/to/report.msg --folder Inbox --column Account --column Amount --occurrence 0

Use this content structure in examples/README.md:

~~~markdown
# Examples

The saved-MSG example does not require Outlook. Live extraction and saving
require Windows desktop Outlook because they use the current pywin32 adapter.

## Read a saved MSG offline

    uv run python -m examples.extract_tables_from_msg path/to/report.msg \
      --column Account --column Amount --occurrence 0

## Read a live Outlook message

    uv run python -m examples.extract_live_tables "Daily Report" \
      --folder Inbox --column Account --column Amount --occurrence 0

## Save and replay a live message

    uv run python -m examples.save_and_extract_message "Daily Report" \
      path/to/report.msg --folder Inbox \
      --column Account --column Amount --occurrence 0

source_index is the absolute zero-based HTML table position. occurrence is the
zero-based position after column filtering. They cannot be supplied together.
Saving refuses a destination that already exists at preflight. Concurrent writers
targeting the same new path must coordinate externally.

Results are Pydantic models. If your application already uses pandas, convert
at the boundary you control (pandas is intentionally not a project dependency):

    import pandas

    frame = pandas.DataFrame(table.rows, columns=table.columns)
~~~

Replace the empty root README.md with this concrete section order:

~~~markdown
# Email Extractor

A Python 3.13 library for filtering Outlook email, extracting attachments and
plain-text bodies, and extracting structured HTML tables from live mail or
saved MSG files.

## Why tables use HTML

EmailRecord.body remains Outlook's plain-text Body for backward compatibility.
Tables use HTMLBody because plain text no longer contains th, tr, or td
structure.

## Install

    uv sync --group dev

Saved-MSG reading is Outlook-independent. Live Outlook access and saving use
the Windows pywin32 adapter.

## Table selection

Positions are zero-based. source_index addresses every table in DOM order:

    TableSelector(source_index=1)

occurrence addresses matching tables after required-column filtering:

    TableSelector(
        required_columns=["Account", "Amount"],
        occurrence=0,
    )

Omit occurrence to return all tables containing the required columns.

## Workflows

Live Outlook records must first be cached by filter_emails:

    records = provider.filter_emails([query])[query.name].records
    tables_by_email = provider.extract_tables(records, selector)

Saved MSG files do not require Outlook:

    tables = extract_tables_from_msg("path/to/report.msg", selector)

Save a cached live message and replay it through the standalone reader:

    saved_path = provider.save_message(records[0], "path/to/report.msg")
    tables = extract_tables_from_msg(saved_path, selector)

Both extraction paths return the same ExtractedTable Pydantic model.

## Pydantic first, pandas optional

The library does not require pandas. A caller that already has pandas may opt in:

    import pandas as pd

    table = tables[0]
    frame = pd.DataFrame(table.rows, columns=table.columns)

## Preconditions and errors

Live extraction and saving require records cached by filter_emails on the same
provider. MSG paths must be existing regular .msg files. Saving requires an
existing parent directory and refuses overwrites.

The collision check occurs before Outlook SaveAs; callers with concurrent writers
targeting one new path must coordinate those writers externally.

- EmailNotInCacheError: the record did not come from this provider instance.
- TableExtractionError: live HTML access or parsing failed.
- MessageFileReadError: the saved MSG could not be decoded.
- InvalidSavePathError: the MSG destination is invalid.
- MessageSaveError: Outlook failed to save the cached message.

See examples/README.md for runnable commands.
~~~

- [ ] **Step 7: Run example tests and direct CLI smoke test**

Run:

    uv run pytest tests/examples/test_examples.py -q
    uv run python -m examples.extract_tables_from_msg tests/fixtures/table_message.msg --column Account --column Amount --occurrence 0

Expected: tests pass and output contains the Account/Amount header and A-100/125.50 row without opening Outlook.

- [ ] **Step 8: Commit examples and documentation**

    git add examples README.md tests/examples/test_examples.py
    git commit -m "docs: add email table extraction examples"

---

### Task 8: Public contract audit and whole-suite verification

**Files:**
- Create: tests/test_public_contract.py
- Modify only if verification exposes a concrete defect in files from Tasks 1-7.

**Interfaces:**
- Consumes: every public model, function, provider method, capability protocol, example, and dependency declaration.
- Produces: final regression coverage proving the feature is additive and both source paths agree.

- [ ] **Step 1: Write the final public-contract test**

Create tests/test_public_contract.py:

~~~python
import inspect

from extractors import extract_html_tables, extract_tables_from_msg
from protocols import MessageSavingEmailProvider, TableExtractingEmailProvider
from protocols.provider import EmailProvider
from providers.outlook import OutlookProvider
from schemas import ExtractedTable, TableSelector


def test_original_provider_protocol_has_no_new_capabilities():
    members = set(EmailProvider.__dict__)
    assert "extract_tables" not in members
    assert "save_message" not in members


def test_new_capability_signatures_match_outlook_provider():
    assert inspect.signature(TableExtractingEmailProvider.extract_tables) == inspect.signature(
        OutlookProvider.extract_tables
    )
    assert inspect.signature(MessageSavingEmailProvider.save_message) == inspect.signature(
        OutlookProvider.save_message
    )


def test_new_public_exports_are_importable():
    assert callable(extract_html_tables)
    assert callable(extract_tables_from_msg)
    assert TableSelector.__name__ == "TableSelector"
    assert ExtractedTable.__name__ == "ExtractedTable"
~~~

These signature assertions are the runtime regression guard for the two structural
protocols. Keep the protocols as ordinary static typing contracts; do not add a new
`@runtime_checkable` API solely for this test.

- [ ] **Step 2: Run the entire suite**

Run:

    uv run pytest -q

Expected: every test passes; Outlook is not launched because provider tests use object.__new__ and fake COM objects.

- [ ] **Step 3: Verify packaging, imports, bytecode, and lock consistency**

Run:

    uv lock --check
    uv run python -m compileall -q exceptions.py schemas protocols providers extractors examples
    uv run python -c "from extractors import extract_tables_from_msg; import sys; assert 'win32com' not in sys.modules; assert 'providers.outlook' not in sys.modules"
    git diff --check

Expected: all commands exit 0, and importing the offline reader loads neither win32com nor the Outlook provider.

- [ ] **Step 4: Re-run the real offline path**

Run:

    uv run python -m examples.extract_tables_from_msg tests/fixtures/table_message.msg --source-index 1

Expected: output contains source_index=1, Account | Amount, and A-100 | 125.50.

- [ ] **Step 5: Review the final diff for compatibility scope**

Run:

    git status --short
    git diff --stat HEAD~7
    git diff HEAD~7 -- protocols/provider.py schemas/filter.py schemas/result.py

Expected: after the seven preceding task commits, HEAD~7 is the committed plan
baseline; protocols/provider.py, schemas/filter.py, and schemas/result.py have no
feature-driven changes, and only files listed in this plan changed.

- [ ] **Step 6: Commit final contract coverage**

    git add tests/test_public_contract.py
    git commit -m "test: verify table source public contracts"

## Final Acceptance Checklist

- [ ] A live cached Outlook message with HTML tables returns selected Pydantic tables.
- [ ] A saved synthetic MSG returns the same table columns/rows without importing win32com.
- [ ] Saving uses Unicode MSG type 9 and refuses an existing destination.
- [ ] source_index is absolute DOM order and occurrence is after column filtering.
- [ ] Nested tables and spans remain deterministic and rectangular.
- [ ] Existing Pydantic fields, EmailProvider, body filtering, and attachments remain unchanged.
- [ ] examples/ documents and runs all three requested workflows.
- [ ] uv run pytest -q, uv lock --check, compileall, and git diff --check all pass.
