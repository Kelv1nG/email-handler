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
