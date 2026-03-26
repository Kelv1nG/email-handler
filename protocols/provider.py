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

    def search_attachments(
        self,
        email_records: list[EmailRecord],
        attachment_names: list[str] | None = None,
        exact_match: bool = False,
    ) -> list[EmailRecord]:
        """Search for emails containing specific attachments.

        Args:
            email_records: List of EmailRecord objects to search through.
            attachment_names: List of attachment names to search for.
                            If None or empty, no filtering is applied.
            exact_match: When True match the full attachment filename exactly.
                        When False match substring (case-insensitive).

        Returns:
            A list of EmailRecord objects that have matching attachments.
        """
        ...

    def save_attachments(
        self,
        email_records: list[EmailRecord],
        save_path: str,
    ) -> dict[str, list[str]]:
        """Save attachments from email records to disk.

        Args:
            email_records: List of EmailRecord objects to extract attachments from.
            save_path: Directory path where attachments will be saved.

        Returns:
            A dict mapping email subjects to lists of saved file paths.

        Raises:
            ValueError: If save_path does not exist or is not writable.
        """
        ...

    def get_attachment_content(
        self,
        email_record: EmailRecord,
    ) -> dict[str, bytes]:
        """Get attachment content directly from email without saving to disk.

        Args:
            email_record: EmailRecord to get attachments from.

        Returns:
            A dict mapping attachment filenames to their binary content.

        Raises:
            ValueError: If email record not found in cache.
        """
        ...
