"""Small, dependency-free helpers shared by the runnable examples."""

from schemas.table import ExtractedTable


def print_tables(tables: list[ExtractedTable]) -> None:
    """Print extracted tables in a compact, readable form."""
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
