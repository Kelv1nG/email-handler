"""Main entry point for email extractor."""

from datetime import datetime

from providers import OutlookProvider
from schemas.filter import DateFilter, FolderFilter, KeywordFilter, SearchQuery


def main():
    # Example: Multiple named queries — keyword filters in folder
    provider = OutlookProvider()
    results = provider.filter_emails([
        SearchQuery(
            name="mediopac_monthly",
            filters=[
                FolderFilter(folder_name="Inbox/ORT"),
                KeywordFilter(keywords=["MEDIOPAC - Monthly - Equity Country Currency Exposure"], exact_match=True),
            ]
        ),
        SearchQuery(
            name="mediogl_monthly",
            filters=[
                FolderFilter(folder_name="Inbox/ORT"),
                KeywordFilter(keywords=["MEDIOGL - Monthly - Equity Country Currency Exposure"], exact_match=True)
            ]
        )
    ])

    # Access results by query name
    for query_name, result in results.items():
        print(f"Query '{query_name}': {len(result.records)} emails found")


if __name__ == "__main__":
    main()
