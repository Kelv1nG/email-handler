"""Abstract protocol for email providers."""

from typing import Protocol

from schemas.filter import AttachmentFilter, BodyFilter, SearchQuery
from schemas.result import BodyContent, EmailKey, EmailRecord, ExtractionResult, Filename, QueryName, QueryResult


class EmailProvider(Protocol):
    """Protocol for email provider implementations."""

    def extract_emails(
        self,
        query: SearchQuery,
    ) -> ExtractionResult:
        """Filter emails then optionally extract attachments and body content.

        Args:
            query: SearchQuery with email_filters and optional attachment_filter/body_filter.

        Returns:
            ExtractionResult with emails, attachments dict, and bodies dict.
            On failure, returns an ExtractionResult with error set.
        """
        ...

    def filter_emails(
        self,
        queries: list[SearchQuery],
    ) -> dict[QueryName, QueryResult]:
        """Filter emails using composeable search queries.

        Args:
            queries: List of SearchQuery objects, each with composeable filters.

        Returns:
            A dict mapping query names to QueryResult objects.

        Raises:
            ValueError: If a folder referenced in a FolderFilter does not exist,
                       or if duplicate query names are provided.
        """
        ...

    def filter_body(
        self,
        emails: list[EmailRecord],
        body_filter: BodyFilter,
    ) -> dict[EmailKey, BodyContent]:
        """Filter emails by body content and return matching bodies.

        Args:
            emails: List of EmailRecord objects to search through.
            body_filter: BodyFilter with keywords and match mode.

        Returns:
            Dict mapping 'subject:timestamp' keys to body text for matching emails.
        """
        ...

    def filter_attachments(
        self,
        email_records: list[EmailRecord],
        attachment_filter: AttachmentFilter,
    ) -> dict[Filename, bytes]:
        """Search for emails with specific attachments and return their content.

        Args:
            email_records: List of EmailRecord objects to search through.
            attachment_filter: AttachmentFilter with filters to apply.

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
    ) -> dict[Filename, bytes]:
        """Get attachment content directly from email without saving to disk.

        Args:
            email_record: EmailRecord to get attachments from.

        Returns:
            A dict mapping attachment filenames to their binary content.

        Raises:
            ValueError: If email record not found in cache.
        """
        ...
