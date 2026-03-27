"""Abstract protocol for email providers."""

from typing import Protocol

from schemas.filter import AttachmentQuery, SearchQuery
from schemas.result import EmailRecord, QueryResult


class EmailProvider(Protocol):
    """Protocol for email provider implementations."""

    def filter_emails(
        self,
        queries: list[SearchQuery],
    ) -> dict[str, QueryResult]:
        """Filter emails using composeable search queries.

        Queries targeting the same folder are batched into a single Outlook
        search using the merged date range, then post-filtered per query.

        Args:
            queries: List of SearchQuery objects, each with composeable filters
                    (FolderFilter, KeywordFilter, DateFilter, etc.).

        Returns:
            A dict mapping query names to QueryResult objects, each containing
            the matching EmailRecord objects for that query.

        Raises:
            ValueError: If a folder referenced in a FolderFilter does not exist,
                       or if duplicate query names are provided.
        """
        ...

    def filter_attachments(
        self,
        email_records: list[EmailRecord],
        attachment_query: AttachmentQuery,
    ) -> dict[str, bytes]:
        """Search for emails with specific attachments and return their content.

        Args:
            email_records: List of EmailRecord objects to search through.
            attachment_query: AttachmentQuery object with filters to apply.

        Returns:
            A dict mapping attachment filenames to their binary content.

        Raises:
            ValueError: If email record not found in cache.
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
