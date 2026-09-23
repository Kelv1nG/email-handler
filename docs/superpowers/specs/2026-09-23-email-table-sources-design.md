# Email Table Sources and MSG Persistence Design

## Summary

Add backward-compatible table extraction from two message sources and one
persistence utility:

1. extract tables from live Outlook messages already found by
   `OutlookProvider`;
2. extract tables from saved `.msg` files without Outlook installed or
   running;
3. save a live cached Outlook message as a Unicode `.msg` file.

Both extraction paths feed the same provider-independent HTML table parser and
return the same Pydantic `ExtractedTable` models. The table parser must not know
whether its HTML came from Outlook COM, a `.msg` file, or a future Office 365
adapter.

The implementation also adds an `examples/` directory with runnable examples
for live extraction, saved-file extraction, and save-then-extract usage.

This design supersedes the source and persistence portions of
`2026-09-18-email-table-extraction-design.md`. It retains that design's table
selection, normalization, rectangularization, compatibility, and Pydantic
result semantics.

## Goals

- Preserve every existing public model, serialized shape, protocol, and method
  signature.
- Keep `EmailRecord.body` as Outlook's existing plain-text `message.Body`.
- Read live table structure on demand from `message.HTMLBody`.
- Read `.msg` files through a standalone parser rather than Outlook COM.
- Save a complete cached Outlook item in Unicode `.msg` format without
  changing or sending it.
- Give both sources identical table-selection and result behavior.
- Keep pandas optional and outside the core API.
- Make a future Office 365 adapter able to reuse the HTML parser without
  changing the table contract.

## Non-Goals

- Sending, forwarding, editing, or deleting email.
- Creating a new synthetic Outlook message through the public library API.
- Adding table fields to `SearchQuery`, `EmailRecord`, or `ExtractionResult`.
- Reading raw RFC822/MIME or `.eml` files in this version.
- Implementing an Office 365 adapter now.
- Reconstructing browser-rendered coordinates or claiming true CSS visual
  left-to-right/top-to-bottom ordering.
- Heuristically classifying layout tables versus data tables.
- Fuzzy, substring, regex, alias, or typed-value column matching.
- Returning pandas `DataFrame` objects from the core API.
- Preserving raw cell HTML, CSS, or embedded-image presentation.

## Existing Constraints

The current repository creates `EmailRecord.body` from `message.Body`, which is
plain text. This discards `<table>`, `<tr>`, `<th>`, and `<td>` structure.
`OutlookProvider.filter_emails()` also retains the live COM objects in
`_message_cache`; attachment operations already depend on this cache.

A `.msg` file is an OLE compound message container, not an HTML document. It
may contain HTML, RTF, and plain-text body representations. The standalone
reader must open the container, obtain its HTML-body bytes, and then invoke the
same HTML parser used by live Outlook.

## Architecture

```text
Live Outlook MailItem.HTMLBody (str) ─┐
                                      ├─> extract_html_tables() ─> ExtractedTable[]
Saved .msg htmlBody (bytes) ──────────┘

Live cached Outlook MailItem ─> save_message() ─> Unicode .msg file
```

The boundary is the HTML value, represented as `str | bytes`:

- Outlook COM supplies `str` through `MailItem.HTMLBody`.
- `extract-msg` supplies `bytes | None` through `MessageBase.htmlBody`.
- Beautiful Soup accepts either form; byte input is passed directly so the
  parser can detect its encoding instead of assuming UTF-8.
- A future Office 365 adapter can pass its HTML body directly to the same pure
  function.

Do not add a generic message-source abstraction in this version. Two explicit
adapters plus a pure HTML function keep the interface clear without adding an
unused abstraction. Source-specific entry points are intentionally named so a
path is never confused with an `EmailRecord`.

## Public Table Models

Add `schemas/table.py` and re-export both models from `schemas.__init__`.

### TableSelector

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

- `source_index`, `occurrence`, and an integer `header_row` are zero-based and
  non-negative.
- `source_index` and `occurrence` are mutually exclusive.
- Required column values are stripped, must remain non-empty, and must be
  unique after comparison normalization.
- An empty `required_columns` list means no column filtering.

Selection order:

1. Enumerate every `<table>` in HTML document preorder, including nested and
   layout tables, assigning its absolute `source_index` before filtering.
2. If `source_index` is supplied, consider only that absolute table.
3. Resolve the table's header according to `header_row`.
4. Keep tables containing every normalized name in `required_columns`.
5. If `occurrence` is supplied, return only that zero-based occurrence among
   the column-matching tables.
6. Otherwise return all qualifying tables in increasing `source_index` order.

### ExtractedTable

```python
class ExtractedTable(BaseModel):
    source_index: int = Field(ge=0)
    caption: str | None = None
    columns: list[str]
    rows: list[list[str | None]]
```

Result invariants:

- `source_index` is absolute among every `<table>` element in source order.
- `caption` is cleaned text from the table's own direct caption or `None`.
- `columns` preserves cleaned display case, duplicates, and blank labels.
- Every data row has exactly `len(columns)` values.
- `""` represents a real empty cell; `None` represents structural padding.
- Rows before and including the selected header are not data rows.

## Shared HTML API

Add `extractors/html_tables.py` and re-export this function from
`extractors.__init__`:

```python
def extract_html_tables(
    html: str | bytes,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    """Extract structured tables from an HTML string or byte sequence."""
```

Behavior:

- `None` selector is equivalent to `TableSelector()`.
- `str` and `bytes` are accepted; other inputs raise `TypeError`.
- Byte input is passed directly to Beautiful Soup. The library must not
  unconditionally decode `.msg` HTML as UTF-8.
- Empty HTML, table-free HTML, unmatched columns, unavailable source indices,
  and out-of-range occurrences return `[]`.
- Parsing is best-effort for malformed email HTML and does not execute scripts
  or styles.

Use Beautiful Soup with Python's built-in `html.parser`. Do not introduce a
browser engine, `lxml`, or `html5lib`.

## HTML Normalization

### Table and Row Boundaries

- Discover all tables in document preorder.
- Parse nested tables independently and assign each its own source index.
- Exclude descendant rows, cells, captions, and text belonging to a nested
  table while parsing its parent.
- Include `<tr>` elements under `<thead>`, `<tbody>`, and `<tfoot>`, plus direct
  `<tr>` children, when their closest table ancestor is the current table.
- Read only direct `th` and `td` cells belonging to each logical row.
- Do not return a table with no logical row containing a cell, although it
  still occupies an absolute source index.

### Cell Text and Column Matching

- Decode HTML entities through the parser.
- Exclude script, style, and nested-table text.
- Join descendant text fragments with a single space.
- Strip outer whitespace and collapse internal whitespace runs to one space.
- Preserve original case in returned captions, columns, and cell values.
- Compare required columns by exact equality after whitespace normalization and
  Unicode `casefold()`.
- Do not perform substring or fuzzy matching.

### Header Selection

An explicit integer `header_row` selects that zero-based logical row across the
table's document-order rows. If it is absent or has no cells, that table does
not qualify.

For `header_row="auto"`:

1. Use the last non-empty row in `<thead>` when present.
2. Otherwise use the first logical row containing a direct `<th>` cell.
3. Otherwise use the first non-empty logical row.
4. If no row qualifies, do not return the table.

Only non-empty rows after the chosen header become data rows.

### Rowspan, Colspan, and Ragged Rows

- Treat missing, non-integer, zero, and negative spans as `1`.
- Repeat a `colspan=N` cell value across all `N` logical columns.
- Carry a `rowspan=N` value through the same logical column for the following
  `N - 1` rows.
- Place new malformed-overlap cells in the next unoccupied logical column.
- Truncate a rowspan extending beyond the final source row.
- Set output width to the largest expanded width of header or data rows.
- Pad short headers with `""` and short data rows with `None`.
- Apply required-column matching after text and span normalization.

## Case 1: Live Outlook Extraction

Add this method only to `OutlookProvider`:

```python
def extract_tables(
    self,
    email_records: list[EmailRecord],
    selector: TableSelector | None = None,
) -> dict[EmailKey, list[ExtractedTable]]:
    """Extract selected tables from cached live Outlook messages."""
```

Behavior:

- Records must have been returned by `filter_emails()` on the same provider
  instance.
- Retrieve each live message with the existing private cache key.
- Read `message.HTMLBody` only when this method is called.
- Include every unique input `EmailKey`; a message with empty HTML or no match
  maps to `[]`.
- Preserve input dictionary insertion order and increasing table source order.
- Reject colliding input result keys with `DuplicateEmailKeysError` rather than
  silently overwriting.
- Raise `EmailNotInCacheError` for a missing cache entry.
- Wrap COM access and unexpected parser failures in `TableExtractionError`,
  chained from the original exception and identifying the email.
- Abort the batch on the first raised failure; do not return partial results.

Add a separate `TableExtractingEmailProvider` protocol in
`protocols/table_provider.py` and export it from `protocols.__init__`. Do not
add the method to the existing `EmailProvider` protocol.

## Case 2: Standalone MSG Extraction

Add `extractors/msg_tables.py` and export:

```python
def extract_tables_from_msg(
    path: str | Path,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    """Extract selected tables from one saved Outlook MSG file."""
```

Behavior:

- Require an existing regular file with a case-insensitive `.msg` suffix.
- Open it with `extract_msg.openMsg(path)` as a context manager so OLE resources
  are always closed.
- Read `message.htmlBody`, whose contract is `bytes | None`.
- Pass returned bytes directly into `extract_html_tables()`.
- Return `[]` if the body is `None`/empty or if no table matches.
- A body synthesized by `extract-msg` from RTF or plain text is accepted, but
  table extraction succeeds only if actual `<table>` markup is present.
- Do not import `win32com`, launch Outlook, or require Outlook to be installed.
- Wrap invalid/corrupt/unsupported `.msg` files, file access failures, body
  access failures, and unexpected parser failures in
  `MessageFileReadError`, chained from the original exception.

Validate the path in this order: existence, regular-file status, then suffix.
Path validation failures use:

- `FileNotFoundError` for a missing path;
- `ValueError` for a non-`.msg` suffix;
- `IsADirectoryError` when the path is not a regular file.

## Additional Utility: Save a Live Message as MSG

Add this Outlook-specific method:

```python
def save_message(
    self,
    email_record: EmailRecord,
    save_path: str | Path,
) -> Path:
    """Save one cached live Outlook message as a Unicode MSG file."""
```

Behavior:

- The record must be cached by the current `OutlookProvider`, matching the
  precondition used by attachment content access.
- `save_path` is a complete destination filename, not a directory.
- Require a case-insensitive `.msg` suffix.
- Require the destination parent to exist and be a directory.
- Refuse to overwrite an existing path and raise `FileExistsError`.
- Call Outlook `MailItem.SaveAs(str(destination), 9)`. Constant `9` is
  `olMSGUnicode`, preserving the Outlook Unicode `.msg` representation.
- Return the destination `Path` after a successful call.
- Raise `EmailNotInCacheError` when the record is not cached.
- Use the existing `InvalidSavePathError` for an invalid parent or suffix.
- Wrap COM save failures in `MessageSaveError`, chained from the original
  exception.
- Do not mutate, send, move, mark read, or delete the live message.

Add a separate `MessageSavingEmailProvider` protocol in
`protocols/message_saver.py` and export it from `protocols.__init__`. Do not add
`save_message()` to `EmailProvider`, preserving structural compatibility for
third-party providers.

The Microsoft API contract is `MailItem.SaveAs(Path, Type)`, and
`olMSGUnicode` has value `9`.

## Exceptions

Add to `exceptions.py`:

```python
class DuplicateEmailKeysError(ValueError):
    """Raised when input records would overwrite the same result key."""


class TableExtractionError(RuntimeError):
    """Raised when live HTML access or table parsing fails unexpectedly."""


class MessageFileReadError(IOError):
    """Raised when a saved MSG file cannot be decoded or parsed."""


class MessageSaveError(IOError):
    """Raised when Outlook cannot save a cached message as MSG."""
```

Normal empty outcomes are not exceptions:

- empty HTML;
- no table elements;
- tables without a usable header;
- unmatched columns;
- unavailable absolute index;
- out-of-range filtered occurrence.

Selector model validation remains Pydantic validation. Live cache, file access,
COM, and parsing failures use the explicit domain errors described above.
The pure `extract_html_tables()` function raises `TableExtractionError`,
chained from the underlying exception, for unexpected HTML-parser failures.

## Dependencies

Runtime dependencies:

- retain `pywin32>=311` on Windows through a `sys_platform == "win32"`
  dependency marker, allowing the standalone reader to install on non-Windows
  systems without `pywin32`;
- retain `pydantic` for public models;
- add `beautifulsoup4` for HTML parsing;
- add `extract-msg>=0.56,<0.57` for standalone `.msg` parsing.

Development dependencies:

- add `pytest` through a `dev` dependency group.

Do not add pandas. Update `uv.lock` so it agrees with `pyproject.toml`, including
the existing Pydantic declaration currently absent from the lock.

## Examples Directory

Add an importable `examples/` package with:

- `examples/README.md` — prerequisites and commands;
- `examples/extract_tables_from_msg.py` — CLI taking a `.msg` path and optional
  required columns, demonstrating offline extraction;
- `examples/extract_live_tables.py` — live Outlook search followed by table
  extraction;
- `examples/save_and_extract_message.py` — save one cached live message as
  `.msg`, then reopen it through the standalone reader and print equivalent
  table results.

Every script must:

- use `if __name__ == "__main__"`;
- be runnable from the repository root with `uv run python -m examples.<name>`;
- avoid hard-coded personal paths, mailbox names, addresses, or subjects;
- use `argparse` for required runtime values;
- print source index, columns, and rows without requiring pandas;
- state when Outlook is required and when it is not;
- avoid sending, changing, or deleting messages.

Do not commit a real user email to `examples/`.

## Testing

### Selector and Pure Parser Tests

Verify:

- selector defaults and validation;
- absolute source indexing across top-level and nested tables;
- column filtering followed by zero-based occurrence;
- automatic and explicit headers;
- nested-content exclusion;
- captions, whitespace, entities, and case-insensitive exact header matching;
- combined rowspan/colspan expansion;
- ragged rows, actual empty cells, and structural padding;
- malformed, empty, and table-free `str` and `bytes` HTML;
- identical results for equivalent `str` and encoded `bytes` HTML.

### Live Outlook Adapter Tests

Use fake cached COM objects; unit tests must not launch Outlook. Verify:

- live `HTMLBody` is parsed without changing `EmailRecord.body`;
- input and table order;
- empty results remain keyed;
- cache misses, duplicate result keys, COM access failures, and parser failures;
- existing filtering, body matching, and attachment behavior remain unchanged.

### MSG Reader Tests

Commit one small, synthetic, non-sensitive `.msg` fixture containing two HTML
tables. The fixture may be generated once with Outlook during development, but
normal tests must read it without Outlook.

Verify:

- the standalone reader obtains HTML and produces the expected Pydantic tables;
- position, column, and occurrence selectors match pure-parser behavior;
- Outlook/`win32com` is not imported by the `.msg` reader module;
- resources close on success and error;
- missing paths, directories, incorrect suffixes, corrupt files, empty bodies,
  and parser failures follow the specified semantics.

### Save Utility Tests

Use a fake cached MailItem and temporary directory; tests must not launch
Outlook. Verify:

- `SaveAs` receives the destination string and `9`;
- a successful save returns the destination `Path`;
- invalid suffix/parent, existing destination, and cache miss errors;
- COM failures become chained `MessageSaveError`;
- no unrelated MailItem mutation methods are called.

### Compatibility and Example Tests

Verify:

- representative `EmailRecord`, `SearchQuery`, and `ExtractionResult`
  `model_dump()` results contain exactly their pre-feature fields;
- `EmailProvider` has no new required method;
- both new capability protocols are structurally satisfied by
  `OutlookProvider`;
- public package exports work;
- every example module imports and compiles without executing its CLI body;
- the offline example runs against the synthetic fixture without Outlook.

## Documentation

Update the root `README.md` to explain:

- why plain `EmailRecord.body` cannot preserve tables;
- live versus `.msg` extraction;
- which operations require Outlook;
- the save-then-replay workflow;
- absolute DOM position versus filtered occurrence;
- cache and file preconditions;
- Pydantic-first results and optional caller-owned pandas conversion.

The runnable `examples/README.md` is the detailed usage guide.

## Backward Compatibility

The change is additive:

- no field is added to an existing Pydantic model;
- no current serialized shape changes;
- no existing signature changes;
- `EmailProvider` remains unchanged;
- existing searches continue reading only plain `message.Body` and incur no
  HTML or `.msg` parsing unless a new API is called;
- existing attachment and body behavior remains unchanged;
- Outlook-free `.msg` extraction lives in a module that does not import
  `win32com`;
- non-Windows installation does not attempt to install `pywin32`;
- pandas remains absent from runtime dependencies.

Adding `extract-msg` and Beautiful Soup changes the installed dependency set but
does not change existing call behavior.

## Alternatives Considered

### Store HTML on EmailRecord

Rejected because adding an `html_body` field changes model serialization and
retains potentially large HTML for every search even when tables are not used.

### Reopen MSG Through Outlook COM

Rejected because it would require Outlook for offline replay and would preserve
the coupling that blocks a future Office 365 source.

### One Overloaded Extract Function

Rejected because accepting HTML, paths, and email records in one function makes
runtime dispatch and error behavior ambiguous. Explicit source adapters feeding
one pure HTML parser are easier to understand and test.

### Generic MessageSource Protocol Now

Deferred because there are only two concrete inputs today. The pure
`extract_html_tables()` boundary already provides the extension seam needed by
a future Office 365 adapter.

## References

- Microsoft `MailItem.SaveAs`:
  <https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.saveas>
- Microsoft `OlSaveAsType.olMSGUnicode = 9`:
  <https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.olsaveastype?view=outlook-pia>
- `extract-msg` `openMsg` API:
  <https://msg-extractor.readthedocs.io/en/latest/extract_msg/open_msg.html>
- `extract-msg` HTML-body contract:
  <https://msg-extractor.readthedocs.io/en/latest/extract_msg/msg_classes/message_base.html>
