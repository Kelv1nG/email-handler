"""Extract selected HTML tables from live Outlook messages."""

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
