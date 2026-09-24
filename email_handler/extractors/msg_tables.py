"""Extract HTML tables from saved Outlook MSG files without Outlook."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

import extract_msg

from email_handler.exceptions import MessageFileReadError
from email_handler.schemas.table import ExtractedTable, TableSelector

from .html_tables import extract_html_tables

if TYPE_CHECKING:
    from polars import DataFrame as PolarsDataFrame
else:
    PolarsDataFrame = Any


def _unique_column_names(columns: list[str]) -> list[str]:
    original_names = set(columns)
    used: set[str] = set()
    next_suffix: dict[str, int] = {}
    unique: list[str] = []

    for column in columns:
        candidate = column
        if candidate in used:
            suffix = next_suffix.get(column, 2)
            candidate = f"{column}_{suffix}"
            while candidate in original_names or candidate in used:
                suffix += 1
                candidate = f"{column}_{suffix}"
            next_suffix[column] = suffix + 1
        used.add(candidate)
        unique.append(candidate)

    return unique


def extracted_table_to_dataframe(table: ExtractedTable) -> PolarsDataFrame:
    """Convert an extracted table to a row-oriented Polars DataFrame."""
    try:
        import polars as pl
    except ModuleNotFoundError as exc:
        if exc.name != "polars":
            raise
        raise ModuleNotFoundError(
            "Polars is required for DataFrame conversion; install "
            "'email-extractor[polars]'."
        ) from exc

    schema = [
        (column, pl.String) for column in _unique_column_names(table.columns)
    ]
    return pl.DataFrame(table.rows, schema=schema, orient="row")


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
