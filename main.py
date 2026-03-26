"""Main entry point for email extractor."""

from datetime import datetime

from providers import OutlookProvider
from schemas.filter import DateFilter, FolderFilter, KeywordFilter, SearchQuery


def main():
    """Example: Extract emails from Outlook using composeable search queries."""
    provider = OutlookProvider()

    # Example 1: Single query — keyword filter in Inbox
    results = provider.filter_emails([
        SearchQuery(filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["rivalry"]),
        ])
    ])
    for query_result in results:
        print(f"Found {len(query_result.records)} emails for query {query_result.query.filters}")
        for email in query_result.records:
            print(f"  - {email.subject} from {email.sender}")

    # Example 2: Multiple queries — same folder batched into one Outlook call
    # results = provider.filter_emails([
    #     SearchQuery(filters=[
    #         FolderFilter(folder_name="Inbox"),
    #         KeywordFilter(keywords=["Budget", "Finance"]),
    #         DateFilter(date_from=datetime(2026, 1, 1)),
    #     ]),
    #     SearchQuery(filters=[
    #         FolderFilter(folder_name="Inbox"),
    #         KeywordFilter(keywords=["Project Update"]),
    #         DateFilter(date_from=datetime(2026, 3, 1)),
    #     ]),
    # ])

    # Example 3: Multiple folders — separate Outlook calls
    # results = provider.filter_emails([
    #     SearchQuery(filters=[
    #         FolderFilter(folder_name="Inbox"),
    #         KeywordFilter(keywords=["rivalry"]),
    #     ]),
    #     SearchQuery(filters=[
    #         FolderFilter(folder_name="Inbox/Archive"),
    #         KeywordFilter(keywords=["rivalry"]),
    #     ]),
    # ])


if __name__ == "__main__":
    main()
