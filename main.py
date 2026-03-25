"""Main entry point for email extractor."""

from datetime import datetime

from providers import OutlookProvider
from schemas import EmailRecord


def main():
    """Example: Extract emails from Outlook with keyword filter."""
    provider = OutlookProvider()

    # Example 1: Filter by keyword
    results = provider.filter_emails(keywords=["rivalry"])
    print(f"Found {len(results)} emails with keyword 'rivalry'")
    for email in results:
        print(f"  - {email.subject} from {email.sender}")

    # Example 2: Filter by date range
    # from_date = datetime(2026, 1, 1)
    # to_date = datetime(2026, 3, 26)
    # results = provider.filter_emails(date_from=from_date, date_to=to_date)
    # print(f"Found {len(results)} emails in date range")

    # Example 3: Exact subject match
    # results = provider.filter_emails(keywords=["Budget Review"], exact_match=True)
    # print(f"Found {len(results)} emails with exact subject")

    # Print results as JSON
    for email in results:
        print(email.model_dump_json(indent=2))


if __name__ == "__main__":
    main()
