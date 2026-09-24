"""Save a live Outlook message and replay its tables offline."""

from argparse import ArgumentParser
from pathlib import Path

from email_handler.extractors import extract_tables_from_msg
from email_handler.schemas import FolderFilter, KeywordFilter, SearchQuery, TableSelector

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

    # Saving requires the current Windows Outlook adapter.
    from email_handler.providers import OutlookProvider

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
