# Email HTML Table Extraction Design

> **Superseded:** The source and persistence portions of this design are
> replaced by
> `docs/superpowers/specs/2026-09-23-email-table-sources-design.md`, which adds
> standalone `.msg` reading, live-message `.msg` saving, and runnable examples.
> The table selector, normalization, and result semantics are retained there.

## Summary

Add an opt-in, backward-compatible API for extracting structured tables from the HTML body of Outlook emails. Existing email searching, plain-text body filtering, attachment extraction, Pydantic models, result serialization, and provider protocol remain unchanged.

The feature adds:

- dependency-neutral Pydantic models for table selection and extracted table data;
- a pure HTML-to-table extraction function that can be tested without Outlook;
- an `OutlookProvider.extract_tables()` capability that reads `HTMLBody` from messages already cached by `filter_emails()`;
- a separate capability protocol instead of changing `EmailProvider`;
- deterministic selection by absolute HTML source position, required column names, and occurrence among matching tables.

Pandas conversion, MIME parsing, visual browser-layout reconstruction, and automatic table extraction through `SearchQuery` or `ExtractionResult` are outside this change.

## Existing Constraints

The repository currently constructs `EmailRecord.body` from Outlook's `message.Body`, which is plain text. Table structure is therefore unavailable from `EmailRecord.body`. Outlook messages are retained in `OutlookProvider._message_cache` after `filter_emails()`, and attachment extraction already relies on that cache for access to data not stored on `EmailRecord`.

The existing public contracts must remain stable:

- `EmailRecord.body` remains plain text with type `str`.
- `SearchQuery`, `BodyFilter`, `ExtractionResult`, and all their serialized fields remain unchanged.
- Existing method signatures and behavior remain unchanged.
- `EmailProvider` remains unchanged so third-party structural implementations do not acquire a new required method.
- HTML is read only when the caller explicitly requests table extraction, so current searches do not gain additional COM calls or HTML parsing cost.

## Public API

### TableSelector

Add `schemas/table.py` with this public request model:

```python
from typing import Literal

from pydantic import BaseModel, Field


class TableSelector(BaseModel):
    source_index: int | None = Field(default=None, ge=0)
    required_columns: list[str] = Field(default_factory=list)
    occurrence: int | None = Field(default=None, ge=0)
    header_row: int | Literal["auto"] = "auto"
```

Validation rules:

- `source_index`, `occurrence`, and an integer `header_row` are zero-based and non-negative.
- `source_index` and `occurrence` are mutually exclusive because they describe different index spaces.
- Required column values are stripped, must remain non-empty, and must be unique after column-name normalization.
- An empty `required_columns` list means no column filtering.

Selection rules are applied in this order:

1. Enumerate every HTML `<table>` element in document preorder, including nested and layout tables, and assign its absolute `source_index` before any parsing or filtering.
2. If `source_index` is present, consider only that absolute table. A missing index returns no match.
3. Determine the table header according to `header_row`.
4. Keep only tables containing every normalized name in `required_columns`.
5. If `occurrence` is present, return only that zero-based occurrence among the remaining tables. If it is out of range, return no match.
6. Without `occurrence`, return every remaining table in increasing `source_index` order.

Examples:

```python
# The second <table> element in HTML source order.
TableSelector(source_index=1)

# The first table, in source order, containing both required columns.
TableSelector(
    required_columns=["Account", "Amount"],
    occurrence=0,
)

# Every table containing an Account column.
TableSelector(required_columns=["Account"])
```

### ExtractedTable

Add this public result model to `schemas/table.py`:

```python
class ExtractedTable(BaseModel):
    source_index: int = Field(ge=0)
    caption: str | None = None
    columns: list[str]
    rows: list[list[str | None]]
```

Result invariants:

- `source_index` is the table's absolute index among all `<table>` elements in the source HTML.
- `caption` contains the cleaned text of the table's own direct caption or `None`.
- `columns` contains cleaned display text from the selected header row. Duplicate and blank labels are preserved because inventing names would lose source fidelity.
- Every row has exactly `len(columns)` cells.
- A real but empty HTML cell is represented by `""`; `None` represents structural padding for a ragged row.
- Header rows and rows preceding the selected header are not included in `rows`.
- Extracted tables are returned in increasing `source_index` order.

`TableSelector` and `ExtractedTable` are re-exported from `schemas.__init__`.

### Pure HTML API

Add an `extractors` package with this public, provider-independent function:

```python
def extract_html_tables(
    html: str,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    ...
```

`None` is equivalent to `TableSelector()` and returns every non-empty parseable table. Empty HTML, HTML without tables, an unmatched column selector, an unavailable absolute index, or an out-of-range occurrence returns `[]`.

Use Beautiful Soup with Python's built-in `html.parser`. Parsing must be best-effort for malformed email HTML, must not evaluate scripts or styles, and must not require pandas or a browser engine.

### Outlook Provider API

Add this method to `OutlookProvider` without changing existing methods:

```python
def extract_tables(
    self,
    email_records: list[EmailRecord],
    selector: TableSelector | None = None,
) -> dict[EmailKey, list[ExtractedTable]]:
    ...
```

Usage:

```python
records = provider.filter_emails([query])[query.name].records
tables_by_email = provider.extract_tables(
    records,
    TableSelector(
        required_columns=["Account", "Amount"],
        occurrence=0,
    ),
)
```

Provider behavior:

- The records must originate from `filter_emails()` on the same `OutlookProvider` instance so their Outlook messages are available in `_message_cache`.
- The provider obtains each cached message with the existing private cache key and reads `message.HTMLBody` only during this method.
- Every unique input `EmailKey` is present in the returned dictionary. An email with empty HTML or no matching tables maps to `[]`.
- Dictionary insertion order follows input email order. Each value is ordered by increasing `source_index`.
- Duplicate input `EmailKey` values raise `DuplicateEmailKeysError` instead of silently overwriting a result. This addresses the existing user-facing key's known `subject:received_time` collision risk without changing existing APIs.
- A record absent from `_message_cache` raises the existing `EmailNotInCacheError`.
- A COM property-access failure or unexpected parser failure raises `TableExtractionError`, chained from the original exception and identifying the affected email.
- The operation is atomic from the caller's perspective: the first domain/runtime failure aborts the call and no partial result is returned.

Add `DuplicateEmailKeysError(ValueError)` and `TableExtractionError(RuntimeError)` to `exceptions.py`.

### Capability Protocol

Add `protocols/table_provider.py`:

```python
class TableExtractingEmailProvider(Protocol):
    def extract_tables(
        self,
        email_records: list[EmailRecord],
        selector: TableSelector | None = None,
    ) -> dict[EmailKey, list[ExtractedTable]]:
        ...
```

Re-export it from `protocols.__init__`. Do not inherit it from or add its method to `EmailProvider`; callers that require both capabilities can type them independently, while existing provider implementations continue satisfying `EmailProvider` unchanged.

## HTML Normalization

### Table and Row Boundaries

- Tables are discovered with document-order traversal of all `<table>` elements.
- A nested table is assigned its own `source_index` and parsed independently.
- When parsing an outer table, descendant rows, cells, captions, and text belonging to a nested table are excluded. This prevents child-table values from being duplicated into the parent.
- Logical rows include `<tr>` elements under `<thead>`, `<tbody>`, and `<tfoot>`, plus direct `<tr>` children, in document order, provided their nearest table ancestor is the table being parsed.
- Only direct `th` and `td` cells belonging to each logical row are considered. Other markup contributes text only when it is inside one of those cells.
- Tables with no logical row containing a cell are not returned, though they retain their place in absolute source indexing.

### Cell Text

For display values and matching:

- HTML entities are decoded by the HTML parser.
- Script, style, and nested-table content are excluded.
- Descendant text fragments are joined with a single space.
- Leading/trailing whitespace is removed and internal whitespace runs are collapsed to one space.
- Original character case is preserved in `caption`, `columns`, and `rows`.

Column matching uses an additional comparison-only normalization: cleaned text is Unicode-case-folded. Matching is exact after normalization; substring, regex, and fuzzy column matching are not part of this version.

### Header Selection

An explicit integer `header_row` selects that zero-based logical row across the table's document-order rows. If it does not exist or contains no cells, that table is not a match.

For `header_row="auto"`:

1. If a `<thead>` has non-empty rows, use its last non-empty logical row. This treats earlier rows as grouping/title rows.
2. Otherwise, use the first logical row containing at least one direct `<th>` cell.
3. Otherwise, use the first non-empty logical row.
4. If no such row exists, the table is not returned.

Only rows after the selected header row become data rows. Empty data rows are discarded.

### Rowspan, Colspan, and Ragged Rows

- Missing, non-integer, zero, and negative span values are treated as `1`.
- `colspan=N` repeats the cell's cleaned value across all `N` logical columns.
- `rowspan=N` repeats the cell's cleaned value into the same logical column for the following `N - 1` rows.
- Active rowspans occupy their columns before new cells in a subsequent row are placed. If malformed HTML would overlap an occupied position, the new cell is placed in the next free logical column.
- A rowspan extending beyond the final source row is silently truncated.
- The output width is the maximum expanded width of the header and retained data rows.
- A short header is padded with `""`. Short data rows are padded with `None`.

Column matching occurs after text normalization and span expansion.

## Dependencies and Packaging

- Add `beautifulsoup4` as a runtime dependency.
- Add `pytest` as a development/test dependency using the project's package-management convention.
- Do not add pandas, lxml, html5lib, a browser engine, or MIME-processing dependencies.
- Keep the repository's existing top-level package layout; this feature does not restructure the project.

## Error Semantics

Normal successful empty results:

- empty or missing `HTMLBody` value;
- no `<table>` elements;
- only empty/unparseable table elements;
- no table containing all required columns;
- unavailable `source_index`;
- out-of-range filtered `occurrence`;
- explicit `header_row` unavailable for a particular table.

Raised errors:

- Pydantic validation error for a negative index, blank/duplicate required column, invalid `header_row`, or simultaneous `source_index` and `occurrence`;
- `DuplicateEmailKeysError` for colliding input result keys;
- `EmailNotInCacheError` for a record not cached by the current provider instance;
- `TableExtractionError` for COM access failures, non-string non-null `HTMLBody` values, or unexpected parser failures.

The standalone pure HTML function raises `TypeError` when `html` is not a string. It does not convert arbitrary objects to strings.

## Compatibility Guarantees

The change is additive:

- No field is added to existing Pydantic models, so their `model_dump()`, JSON, and generated schemas remain stable.
- No existing function or method signature changes.
- No current protocol gains a new required method.
- Existing code continues using Outlook's plain-text `Body`; it does not read or retain `HTMLBody` unless the new method is called.
- Existing result-key behavior is unchanged outside the new method.
- Importing and using existing modules does not require pandas.

Backward compatibility tests will capture representative `model_dump()` output for `EmailRecord`, `SearchQuery`, and `ExtractionResult` and assert that the table feature introduces no additional fields.

## Test Strategy

### Selector Model Tests

Verify:

- defaults select all tables;
- negative indices fail validation;
- simultaneous absolute and filtered positions fail validation;
- blank and normalization-equivalent duplicate required columns fail validation;
- explicit non-negative header rows are accepted.

### Pure Parser Tests

Use small inline HTML fixtures to verify:

- top-level and nested tables receive absolute preorder indices;
- nested text is absent from the parent result;
- `source_index` counts tables that fail column filtering;
- required columns are all-required, whitespace-normalized, entity-decoded, and case-insensitive;
- `occurrence` is counted only after column filtering and is zero-based;
- omitting `occurrence` returns all qualifying tables;
- automatic `<thead>`, `<th>`, and headerless-table detection;
- explicit header selection and exclusion of prior/header rows from data;
- blank and duplicate headings remain present;
- combined rowspan and colspan expansion;
- empty cells versus structural padding;
- ragged and malformed rows remain rectangular;
- caption extraction;
- empty, table-free, and malformed HTML behavior;
- result ordering by `source_index`.

### Outlook Provider Tests

Use fake cached Outlook message objects; tests must not launch Outlook. Verify:

- `HTMLBody` is parsed and plain `EmailRecord.body` is untouched;
- input order and per-email table order are retained;
- every unique input email key appears, including empty matches;
- a missing cached message raises `EmailNotInCacheError`;
- duplicate result keys raise `DuplicateEmailKeysError`;
- an `HTMLBody` property failure and parser failure become chained `TableExtractionError` exceptions;
- a failure aborts rather than returning partial results;
- existing filtering and attachment behavior remains unchanged.

### Protocol and Compatibility Tests

Verify:

- `OutlookProvider` structurally satisfies `TableExtractingEmailProvider`;
- the original `EmailProvider` contract has no new member;
- existing Pydantic dumps contain exactly their pre-feature fields;
- new public models and capability protocol can be imported from their package exports.

## Documentation

Update `README.md` with:

- the required `filter_emails()`-then-`extract_tables()` flow;
- examples for absolute position, columns plus occurrence, and all matching tables;
- the zero-based DOM-order definition;
- the cache precondition and raised errors;
- the fact that table extraction uses Outlook `HTMLBody` while `EmailRecord.body` remains plain text;
- the dependency-neutral result shape and a note that callers may construct pandas objects themselves, without making pandas a package dependency.

## Non-Goals

- Returning pandas `DataFrame` objects from the core API.
- Adding table fields to `SearchQuery`, `EmailRecord`, or `ExtractionResult`.
- Parsing raw RFC822/MIME messages or non-Outlook providers.
- Reconstructing rendered CSS coordinates or claiming true visual left-to-right/top-to-bottom order.
- Classifying layout versus data tables heuristically.
- Fuzzy, substring, regex, alias, or typed-value column matching.
- Inferring numeric, date, boolean, or currency cell types.
- Preserving raw cell HTML or style information.
