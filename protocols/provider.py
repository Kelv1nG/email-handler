"""Abstract protocol for email providers."""

from datetime import datetime
from typing import Protocol

from schemas.email import EmailRecord


class EmailProvider(Protocol):
    """Protocol for email provider implementations."""

    def filter_emails(
        self,
        keywords: list[str] | None = None,
        exact_match: bool = False,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
        folder_name: str = "Inbox",
    ) -> list[EmailRecord]:
        """Filter emails by keywords and date range.

        Args:
            keywords: List of keywords to search for in subject and body.
                     If None or empty, no keyword filter is applied.
            exact_match: When True the keyword must match the full subject/body.
                        When False a substring (case-insensitive) match is used.
            date_from: Only include emails received on or after this datetime.
                      Defaults to None (no lower bound).
            date_to: Only include emails received on or before this datetime.
                    Defaults to None (no upper bound).
            folder_name: Email folder to search in. Defaults to "Inbox".

        Returns:
            A list of EmailRecord objects matching the filters.

        Raises:
            ValueError: If folder_name does not exist or other validation errors.
        """
        ...
