# Examples

Run these modules from the repository root after installing the project.
Reusable application code lives in the `email_handler` package; for example,
`from email_handler.schemas import TableSelector` and
`from email_handler.extractors import extract_tables_from_msg`.

The saved-MSG example does not require Outlook. Live extraction and saving
require Windows desktop Outlook because they use the current pywin32 adapter.

## Read a saved MSG offline

    uv run python -m examples.extract_tables_from_msg path/to/report.msg --column Account --column Amount --occurrence 0

## Read a live Outlook message

    uv run python -m examples.extract_live_tables "Daily Report" --folder Inbox --column Account --column Amount --occurrence 0

## Save and replay a live message

    uv run python -m examples.save_and_extract_message "Daily Report" path/to/report.msg --folder Inbox --column Account --column Amount --occurrence 0

If more than one message has that exact subject, save-and-replay uses the newest
match returned by Outlook.

`source_index` is the absolute zero-based HTML table position. `occurrence` is
the zero-based position after column filtering. They cannot be supplied together.
Saving refuses a destination that already exists at preflight. Concurrent writers
targeting the same new path must coordinate externally.

Results are Pydantic models. Install the optional Polars integration and use the
public conversion helper when you want a Polars DataFrame:

```text
python -m pip install "email-extractor[polars]"
```

```python
from email_handler.extractors import extracted_table_to_dataframe

frame = extracted_table_to_dataframe(table)
```

If your application already uses pandas, you can instead convert at the
boundary you control (pandas is intentionally not a project dependency):

```python
import pandas

frame = pandas.DataFrame(table.rows, columns=table.columns)
```
