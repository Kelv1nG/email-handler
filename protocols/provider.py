"""Abstract protocol for email providers."""

from typing import Protocol

from schemas.filter import SearchQuery
from schemas.result import EmailRecord, QueryResult


class EmailProvider(Protocol):
    """Protocol for email provider implementations."""

    def filter_emails(
        self,
        queries: list[SearchQuery],
    ) -> list[QueryResult]:
        """Filter emails using composeable search queries.

        Queries targeting the same folder are batched into a single Outlook
        search using the merged date range, then post-filtered per query.

        Args:
            queries: List of SearchQuery objects, each with composeable filters
                    (FolderFilter, KeywordFilter, DateFilter, etc.).

        Returns:
            A list of QueryResult objects, one per input query, each containing
            the matching EmailRecord objects for that query.

        Raises:
            ValueError: If a folder referenced in a FolderFilter does not exist.
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
