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


def _to_table(
    grid: HtmlTableGrid, selector: TableSelector
) -> ExtractedTable | None:
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
        value if value is not None else "" for value in header.values
    ] + [""] * (width - len(header.values))
    rows = [
        list(row.values) + [None] * (width - len(row.values)) for row in data
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


def extract_html_tables(
    html: str | bytes,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    """Extract structured tables from an HTML string or byte sequence."""
    if not isinstance(html, (str, bytes)):
        raise TypeError("html must be str or bytes")
    selected = selector or TableSelector()
    try:
        grids = (
            parse_html_grids(html)
            if selected.source_index is None
            else parse_html_grids(html, source_index=selected.source_index)
        )

        tables = [
            table
            for grid in grids
            if (table := _to_table(grid, selected)) is not None
            and _has_required_columns(table, selected.required_columns)
        ]
    except TableExtractionError:
        raise
    except Exception as exc:
        raise TableExtractionError(
            f"Failed to extract HTML tables: {exc}"
        ) from exc

    if selected.occurrence is None:
        return tables
    if selected.occurrence >= len(tables):
        return []
    return [tables[selected.occurrence]]
