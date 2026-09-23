# Email Extractor

A Python 3.13 library for filtering Outlook email, extracting attachments and
plain-text bodies, and extracting structured HTML tables from live mail or
saved MSG files.

## Why tables use HTML

`EmailRecord.body` remains Outlook's plain-text `Body` for backward compatibility.
Tables use `HTMLBody` because plain text no longer contains `th`, `tr`, or `td`
structure.

## Install

    uv sync --group dev

Saved-MSG reading is Outlook-independent. Live Outlook access and saving use
the Windows pywin32 adapter.

## Table selection

Positions are zero-based. `source_index` addresses every table in DOM order:

```python
TableSelector(source_index=1)
```

`occurrence` addresses matching tables after required-column filtering:

```python
TableSelector(
    required_columns=["Account", "Amount"],
    occurrence=0,
)
```

Omit `occurrence` to return all tables containing the required columns.

## Workflows

Live Outlook records must first be cached by `filter_emails`:

```python
records = provider.filter_emails([query])[query.name].records
tables_by_email = provider.extract_tables(records, selector)
```

Saved MSG files do not require Outlook:

```python
tables = extract_tables_from_msg("path/to/report.msg", selector)
```

Save a cached live message and replay it through the standalone reader:

```python
saved_path = provider.save_message(records[0], "path/to/report.msg")
tables = extract_tables_from_msg(saved_path, selector)
```

Both extraction paths return the same `ExtractedTable` Pydantic model.

## Pydantic first, pandas optional

The library does not require pandas. A caller that already has pandas may opt in:

```python
import pandas as pd

table = tables[0]
frame = pd.DataFrame(table.rows, columns=table.columns)
```

## Preconditions and errors

Live extraction and saving require records cached by `filter_emails` on the same
provider. MSG paths must be existing regular `.msg` files. Saving requires an
existing parent directory and refuses overwrites.

The collision check occurs before Outlook `SaveAs`; callers with concurrent writers
targeting one new path must coordinate those writers externally.

- `EmailNotInCacheError`: the record did not come from this provider instance.
- `DuplicateEmailKeysError`: two requested live records map to the same result key.
- `TableExtractionError`: live HTML access or parsing failed.
- `MessageFileReadError`: the saved MSG could not be decoded.
- `InvalidSavePathError`: the MSG destination is invalid.
- `MessageSaveError`: Outlook failed to save the cached message.

Saved-MSG path validation may also raise `FileNotFoundError`, `IsADirectoryError`,
or `ValueError`. Saving to a path that already exists raises `FileExistsError`.

The snippets above focus on the relevant calls. See
[examples/README.md](examples/README.md) for complete runnable commands.
