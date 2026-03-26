"""Outlook email provider implementation."""

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
import tempfile

import win32com.client

from schemas.filter import DateFilter, FolderFilter, KeywordFilter, SearchQuery
from schemas.result import EmailRecord, QueryResult
from utils.dates import format_outlook_date

# Outlook object model constants from pywin32
# Reference: https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/
OL_MAIL_CLASS = 43  # olMail - represents a mail message item


class OutlookProvider:
    """Outlook email provider for filtering and extracting emails.

    Uses pywin32 (win32com) to interact with Microsoft Outlook COM objects.
    """

    def __init__(self):
        """Initialize Outlook provider with pywin32 COM dispatch."""
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self._message_cache: dict[str, Any] = {}  # Cache for Outlook message objects

    # ============================================================================
    # Protocol Methods (EmailProvider Interface)
    # ============================================================================

    def filter_emails(self, queries: list[SearchQuery]) -> list[QueryResult]:
        """Filter Outlook emails using composeable search queries.

        Queries targeting the same folder are batched into a single Outlook
        search using the merged date range, then post-filtered per query.

        Args:
            queries: List of SearchQuery objects with composeable filters.

        Returns:
            A list of QueryResult objects, one per input query.

        Raises:
            ValueError: If a folder referenced in a FolderFilter does not exist.
        """
        folder_groups = self._group_queries_by_folder(queries)
        query_records: dict[int, list[EmailRecord]] = {i: [] for i in range(len(queries))}

        for folder_name, indexed_queries in folder_groups.items():
            merged_date_from, merged_date_to = self._merge_date_ranges(indexed_queries)
            merged_keywords, all_exact_match = self._merge_keywords(indexed_queries)
            
            # Single Outlook search with merged keywords + date range (server-side filtering)
            folder_emails = self._search_folder(
                folder_name, merged_keywords, all_exact_match, merged_date_from, merged_date_to
            )

            # Post-filter each email against each query's specific keyword settings
            for idx, query in indexed_queries:
                matched = self._apply_query_filters(folder_emails, query)
                query_records[idx].extend(matched)

        return [QueryResult(query=queries[i], records=query_records[i]) for i in range(len(queries))]

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
                            If None or empty, all records are returned.
            exact_match: When True match the full attachment filename exactly.
                        When False match substring (case-insensitive).

        Returns:
            A list of EmailRecord objects that have matching attachments.
        """
        if not attachment_names:
            return email_records

        results: list[EmailRecord] = []

        for record in email_records:
            if not record.attachments:
                continue

            for attachment_name in attachment_names:
                for filename in record.attachments:
                    if exact_match:
                        if attachment_name == filename:
                            results.append(record)
                            break
                    else:
                        if attachment_name.lower() in filename.lower():
                            results.append(record)
                            break

                # Break out of attachment_names loop if we found a match
                if results and results[-1] == record:
                    break

        # Remove duplicates while preserving order
        seen = set()
        unique_results = []
        for record in results:
            record_id = id(record)
            if record_id not in seen:
                seen.add(record_id)
                unique_results.append(record)

        return unique_results

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
        save_dir = Path(save_path)

        if not save_dir.exists():
            raise ValueError(f"Save path '{save_path}' does not exist.")

        if not save_dir.is_dir():
            raise ValueError(f"Save path '{save_path}' is not a directory.")

        results: dict[str, list[str]] = {}

        for record in email_records:
            if not record.attachments:
                continue

            # Get the cached Outlook message
            cache_key = self._get_cache_key(record)
            message = self._message_cache.get(cache_key)

            if not message:
                print(f"Warning: Message not found in cache for subject '{record.subject}'")
                continue

            saved_files: list[str] = []

            try:
                for i in range(message.Attachments.Count):
                    attachment = message.Attachments.Item(i + 1)
                    filename = attachment.FileName
                    file_path = save_dir / filename

                    # Save attachment to disk
                    attachment.SaveAsFile(str(file_path))
                    saved_files.append(str(file_path))

            except Exception as e:
                print(f"Error saving attachments for '{record.subject}': {e}")
                continue

            if saved_files:
                results[record.subject] = saved_files

        return results

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
        cache_key = self._get_cache_key(email_record)
        message = self._message_cache.get(cache_key)

        if not message:
            raise ValueError(
                f"Email record for '{email_record.subject}' not found in cache. "
                "Make sure to call filter_emails() first."
            )

        results: dict[str, bytes] = {}

        try:
            for i in range(message.Attachments.Count):
                attachment = message.Attachments.Item(i + 1)
                filename = attachment.FileName

                # Save to temp file, read content, then delete temp file
                with tempfile.TemporaryDirectory() as tmpdir:
                    temp_path = Path(tmpdir) / filename
                    attachment.SaveAsFile(str(temp_path))

                    # Read content into memory
                    with open(temp_path, "rb") as f:
                        content = f.read()
                    results[filename] = content

        except Exception as e:
            raise ValueError(f"Error reading attachments: {e}") from e

        return results

    # ============================================================================
    # Private Helper Methods
    # ============================================================================

    def _group_queries_by_folder(
        self, queries: list[SearchQuery]
    ) -> dict[str, list[tuple[int, SearchQuery]]]:
        """Group queries by their folder name, preserving original index."""
        groups: dict[str, list[tuple[int, SearchQuery]]] = defaultdict(list)
        for i, query in enumerate(queries):
            folder_filter = next((f for f in query.filters if isinstance(f, FolderFilter)), None)
            folder_name = folder_filter.folder_name if folder_filter else "Inbox"
            groups[folder_name].append((i, query))
        return groups

    def _merge_date_ranges(
        self, indexed_queries: list[tuple[int, SearchQuery]]
    ) -> tuple[datetime | None, datetime | None]:
        """Return the widest date range covering all queries in a folder group."""
        date_froms = [
            f.date_from
            for _, q in indexed_queries
            for f in q.filters
            if isinstance(f, DateFilter) and f.date_from is not None
        ]
        date_tos = [
            f.date_to
            for _, q in indexed_queries
            for f in q.filters
            if isinstance(f, DateFilter) and f.date_to is not None
        ]
        return (min(date_froms) if date_froms else None, max(date_tos) if date_tos else None)

    def _merge_keywords(
        self, indexed_queries: list[tuple[int, SearchQuery]]
    ) -> tuple[list[str], bool]:
        """Collect keywords and determine if all are exact match.
        
        Returns:
            Tuple of (keywords list, whether all queries use exact_match).
        """
        keywords = set()
        all_exact_match = True
        for _, q in indexed_queries:
            for f in q.filters:
                if isinstance(f, KeywordFilter):
                    keywords.update(f.keywords)
                    if not f.exact_match:
                        all_exact_match = False
        return list(keywords), all_exact_match

    def _apply_query_filters(
        self, emails: list[EmailRecord], query: SearchQuery
    ) -> list[EmailRecord]:
        """Post-filter a list of emails against a single query's keyword and date filters."""
        keyword_filter = next((f for f in query.filters if isinstance(f, KeywordFilter)), None)
        date_filter = next((f for f in query.filters if isinstance(f, DateFilter)), None)

        results = []
        for email in emails:
            if keyword_filter and not self._matches_keywords(
                email.subject, keyword_filter.keywords, keyword_filter.exact_match
            ):
                continue

            if date_filter:
                # Strip tzinfo for comparison (Outlook returns tz-aware datetimes)
                recv = email.received_time.replace(tzinfo=None)
                if date_filter.date_from and recv < date_filter.date_from.replace(tzinfo=None):
                    continue
                if date_filter.date_to and recv > date_filter.date_to.replace(tzinfo=None):
                    continue

            results.append(email)
        return results

    def _search_folder(
        self,
        folder_name: str,
        keywords: list[str] | None,
        exact_match: bool,
        date_from: datetime | None,
        date_to: datetime | None,
    ) -> list[EmailRecord]:
        """Fetch mail items from a folder with server-side filtering.

        Keywords and date range are applied by Outlook before fetching results.

        Args:
            folder_name: Outlook folder to search.
            keywords: Keywords to search for in subject (OR logic). None for no keyword filter.
            exact_match: If True, match full subject exactly. If False, substring match.
            date_from: Lower bound for ReceivedTime (inclusive). None for no bound.
            date_to: Upper bound for ReceivedTime (inclusive). None for no bound.

        Returns:
            List of EmailRecord objects matching the filters.
        """
        folder = self._get_folder(folder_name)
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)

        # Build server-side restriction: date range always works;
        # keyword exact match uses [Subject] = "..." (also valid server-side);
        # substring keyword search is not supported by Restrict() and must be client-side.
        restrictions = []
        if exact_match and keywords:
            restrictions.append(self._build_keyword_restriction(keywords))
        if date_from is not None or date_to is not None:
            restrictions.append(self._build_date_restriction(date_from, date_to))

        if restrictions:
            combined = " AND ".join(restrictions)
            print(f"Outlook restriction: {combined}")
            messages = messages.Restrict(combined)

        results: list[EmailRecord] = []
        for message in messages:
            if not self._is_mail_item(message):
                continue

            subject: str = message.Subject or ""
            attachments = self._extract_attachments(message)

            email_record = EmailRecord(
                subject=subject,
                sender=message.SenderName or "",
                sender_email=message.SenderEmailAddress or "",
                received_time=message.ReceivedTime,
                body=message.Body or "",
                attachments=attachments,
            )

            cache_key = self._get_cache_key(email_record)
            self._message_cache[cache_key] = message
            results.append(email_record)

        return results

    def _build_keyword_restriction(self, keywords: list[str]) -> str:
        """Build Outlook restriction string for exact subject match (OR logic).

        Note: Outlook's Restrict() does not support substring/contains matching.
        This method is only valid for exact_match=True queries.
        Substring matching must be done client-side in _apply_query_filters.

        Args:
            keywords: List of keywords for exact subject match.

        Returns:
            Outlook restriction string matching any keyword exactly.
        """
        if not keywords:
            return ""
        # Escape double quotes by doubling them
        escaped = [kw.replace('"', '""') for kw in keywords]
        conditions = [f'[Subject] = "{kw}"' for kw in escaped]
        return " OR ".join(conditions)

    def _build_date_restriction(
        self, date_from: datetime | None, date_to: datetime | None
    ) -> str:
        """Build Outlook restriction string for date range.
        
        Args:
            date_from: Start date (inclusive).
            date_to: End date (inclusive).
            
        Returns:
            Outlook restriction string for date range.
        """
        conditions = []
        if date_from is not None:
            formatted = format_outlook_date(date_from)
            conditions.append(f"[ReceivedTime] >= '{formatted}'")
        if date_to is not None:
            formatted = format_outlook_date(date_to)
            conditions.append(f"[ReceivedTime] <= '{formatted}'")
        return " AND ".join(conditions) if conditions else ""

    def _get_folder(self, folder_name: str) -> Any:
        """Get folder by name (supports subfolders).

        Args:
            folder_name: Name of the folder to retrieve. Use format "Inbox" or "Inbox/subfolder"
                        for nested folders. For subfolders only, just pass the subfolder name.

        Returns:
            Outlook folder object.

        Raises:
            ValueError: If folder does not exist.
        """
        inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        # Handle nested folder paths (e.g., "Inbox/ort" or "ort")
        if "/" in folder_name:
            parts = folder_name.split("/")
            folder = inbox
            for part in parts:
                if part.lower() == "inbox":
                    continue
                try:
                    folder = folder.Folders[part]
                except Exception as e:
                    raise ValueError(f"Outlook folder '{folder_name}' not found.") from e
            return folder

        # Handle single folder names
        if folder_name.lower() == "inbox":
            return inbox

        # Try to get subfolder from inbox
        try:
            return inbox.Folders[folder_name]
        except Exception as e:
            # Fallback: try to get from parent folders
            try:
                return inbox.Parent.Folders[folder_name]
            except Exception:
                raise ValueError(f"Outlook folder '{folder_name}' not found.") from e

    def _is_mail_item(self, message: Any) -> bool:
        """Check if message is a mail item (not meeting request, etc.).

        Args:
            message: Outlook message object.

        Returns:
            True if message is a mail item, False otherwise.
        """
        try:
            return message.Class == OL_MAIL_CLASS
        except Exception:
            return False

    def _matches_keywords(
        self, subject: str, keywords: list[str], exact_match: bool
    ) -> bool:
        """Check if subject matches any keyword.

        Args:
            subject: Email subject to check.
            keywords: Keywords to search for.
            exact_match: If True, match full subject; if False, substring match.

        Returns:
            True if subject matches any keyword, False otherwise.
        """
        for kw in keywords:
            if exact_match:
                if kw == subject:
                    return True
            else:
                if kw.lower() in subject.lower():
                    return True
        return False

    def _extract_attachments(self, message: Any) -> list[str]:
        """Extract attachment filenames from message.

        Args:
            message: Outlook message object.

        Returns:
            List of attachment filenames.
        """
        try:
            return [
                message.Attachments.Item(i + 1).FileName
                for i in range(message.Attachments.Count)
            ]
        except Exception:
            return []

    def _get_cache_key(self, record: EmailRecord) -> str:
        """Generate a cache key for an email record.

        Args:
            record: EmailRecord to generate cache key for.

        Returns:
            A unique cache key string.
        """
        return f"{record.sender_email}:{record.received_time.isoformat()}:{record.subject}"
