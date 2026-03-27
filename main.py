"""Main entry point for email extractor."""

from datetime import datetime

from providers import OutlookProvider
from schemas.filter import AttachmentFilter, BodyFilter, DateFilter, FolderFilter, KeywordFilter, SearchQuery


def main():
    provider = OutlookProvider()

    # Filter emails only
    result = provider.extract_emails(SearchQuery(
        name="mediopac",
        email_filters=[
            FolderFilter(folder_name="Inbox/ORT"),
            KeywordFilter(keywords=["MEDIOPAC - Monthly - Equity Country Currency Exposure Report"], exact_match=True),
            DateFilter(date_from=datetime(2026, 1, 1), date_to=datetime(2026, 3, 31)),
        ],
    ))
    print(result)

    # Filter emails + extract pdf attachments + filter body by keyword
    result = provider.extract_emails(SearchQuery(
        name="mediopac",
        email_filters=[
            FolderFilter(folder_name="Inbox/ORT"),
            KeywordFilter(keywords=["MEDIOPAC - Monthly - Equity Country Currency Exposure Report"], exact_match=True),
            DateFilter(date_from=datetime(2026, 1, 1), date_to=datetime(2026, 3, 31)),
        ],
        attachment_filter=AttachmentFilter(filters=[KeywordFilter(keywords=["pdf"])]),
        body_filter=BodyFilter(keywords=["equity exposure"]),
    ))
    print(result)

    if result.error:
        print(f"Error: {result.error}")
        return

    print(f"Emails found: {len(result.emails)}")
    print(f"Attachments: {list(result.attachments.keys())}")
    print(f"Bodies matched: {list(result.bodies.keys())}")


if __name__ == "__main__":
    main()
