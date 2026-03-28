"""Outlook email provider implementation."""

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
import tempfile

import win32com.client

from exceptions import AttachmentReadError, DuplicateQueryNamesError, EmailNotInCacheError, FolderNotFoundError, InvalidSavePathError
from schemas.filter import AttachmentFilter, BodyFilter, DateFilter, FolderFilter, KeywordFilter, SearchQuery
from schemas.result import AttachmentContent, BodyContent, EmailKey, EmailRecord, ExtractionResult, Filename, QueryName, QueryResult
from utils.dates import format_outlook_date

# Outlook object model constants from pywin32
# Reference: https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/
OL_MAIL_CLASS = 43  # olMail - represents a mail message item


class OutlookProvider:
    """Outlook email provider for filtering and extracting emails.

    Uses pywin32 (win32com) to interact with Microsoft Outlook COM objects.
    """

    def __init__(self, temp_dir: str | None = None):
        """Initialize Outlook provider with pywin32 COM dispatch.

        Args:
            temp_dir: Optional directory to use for temporary files when reading
                attachment content. Useful when this module is used as a submodule
                and the default system temp directory causes permission errors.
                Defaults to the system temp directory.
        """
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self._message_cache: dict[str, Any] = {}  # Cache for Outlook message objects
        self._temp_dir = temp_dir

    # ============================================================================
    # Protocol Methods (EmailProvider Interface)
    # ============================================================================

    def filter_emails(self, queries: list[SearchQuery]) -> dict[QueryName, QueryResult]:
        """Filter Outlook emails using composeable search queries.

        Queries targeting the same folder are batched into a single Outlook
        search using the merged date range, then post-filtered per query.

        Args:
            queries: List of SearchQuery objects with composeable filters.

        Returns:
            A dict mapping query names to QueryResult objects, one per input query.

        Raises:
            ValueError: If a folder referenced in a FolderFilter does not exist,
                       or if duplicate query names are provided.
        """
        # Validate unique query names
        query_names = [q.name for q in queries]
        if len(query_names) != len(set(query_names)):
            duplicates = [name for name in set(query_names) if query_names.count(name) > 1]
            raise DuplicateQueryNamesError(f"Duplicate query names found: {duplicates}")
        
        folder_groups = self._group_queries_by_folder(queries)
        query_records: dict[QueryName, list[EmailRecord]] = {q.name: [] for q in queries}

        for folder_name, queries_in_folder in folder_groups.items():
            merged_date_from, merged_date_to = self._merge_date_ranges(queries_in_folder)
            merged_keywords, all_exact_match = self._merge_keywords(queries_in_folder)
            
            # Single Outlook search with merged keywords + date range (server-side filtering)
            folder_emails = self._search_folder(
                folder_name, merged_keywords, all_exact_match, merged_date_from, merged_date_to
            )

            # Post-filter each email against each query's specific keyword settings
            for query in queries_in_folder:
                matched = self._apply_query_filters(folder_emails, query)
                query_records[query.name].extend(matched)

        return {q.name: QueryResult(query=q, records=query_records[q.name]) for q in queries}

    def extract_emails(self, query: SearchQuery) -> ExtractionResult:
        """Filter emails then optionally extract attachments and body content.

        Args:
            query: SearchQuery with email_filters and optional attachment_filter/body_filter.

        Returns:
            ExtractionResult with emails, attachments dict, and bodies dict.
            On failure, returns an ExtractionResult with error set.
        """
        try:
            query_results = self.filter_emails([query])
            filtered_emails = query_results[query.name].records

            attachments: dict[EmailKey, AttachmentContent] = {}
            if query.attachment_filter is not None:
                attachments = self._extract_attachments_by_email(filtered_emails, query.attachment_filter)

            bodies: dict[EmailKey, BodyContent] = {}
            if query.body_filter is not None:
                bodies = self.filter_body(filtered_emails, query.body_filter)

            return ExtractionResult(name=query.name, emails=filtered_emails, attachments=attachments, bodies=bodies)

        except Exception as e:
            return ExtractionResult(name=query.name, error=str(e))

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
        result: dict[EmailKey, BodyContent] = {}
        for email in emails:
            if self._matches_keywords(email.body, body_filter.keywords, body_filter.exact_match):
                result[self._make_email_key(email)] = email.body
        return result

    def filter_attachments(
        self,
        email_records: list[EmailRecord],
        attachment_query: AttachmentFilter
    ) -> dict[Filename, bytes]:
        """Search for emails with specific attachments and return their content.

        Args:
            email_records: List of EmailRecord objects to search through.
            attachment_query: AttachmentQuery object with filters to apply.

        Returns:
            A dict mapping attachment filenames to their binary content.

        Raises:
            ValueError: If email record not found in cache.
        """
        attachment_names, exact_match = self._extract_attachment_filters(attachment_query)
        matching_records = self._search_attachment_records(email_records, attachment_names, exact_match)
        return self._extract_attachment_contents(matching_records, attachment_query)

    # TODO more granular save attachments
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
            raise InvalidSavePathError(f"Save path '{save_path}' does not exist.")

        if not save_dir.is_dir():
            raise InvalidSavePathError(f"Save path '{save_path}' is not a directory.")

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
    ) -> dict[Filename, bytes]:
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
            raise EmailNotInCacheError(
                f"Email record for '{email_record.subject}' not found in cache. "
                "Make sure to call filter_emails() first."
            )

        results: dict[Filename, bytes] = {}

        try:
            for i in range(message.Attachments.Count):
                attachment = message.Attachments.Item(i + 1)
                filename = attachment.FileName

                # Save to temp file, read content, then delete temp file
                with tempfile.TemporaryDirectory(dir=self._temp_dir) as tmpdir:
                    temp_path = Path(tmpdir) / filename
                    attachment.SaveAsFile(str(temp_path))

                    # Read content into memory
                    with open(temp_path, "rb") as f:
                        content = f.read()
                    results[filename] = content

        except Exception as e:
            raise AttachmentReadError(f"Error reading attachments: {e}") from e

        return results

    # ============================================================================
    # Private Helper Methods
    # ============================================================================

    def _group_queries_by_folder(
        self, queries: list[SearchQuery]
    ) -> dict[str, list[SearchQuery]]:
        """Group queries by their folder name."""
        groups: dict[str, list[SearchQuery]] = defaultdict(list)
        for query in queries:
            folder_filter = next((f for f in query.email_filters if isinstance(f, FolderFilter)), None)
            folder_name = folder_filter.folder_name if folder_filter else "Inbox"
            groups[folder_name].append(query)
        return groups

    def _merge_date_ranges(
        self, queries: list[SearchQuery]
    ) -> tuple[datetime | None, datetime | None]:
        """Return the widest date range covering all queries in a folder group."""
        date_froms = [
            f.date_from
            for q in queries
            for f in q.email_filters
            if isinstance(f, DateFilter) and f.date_from is not None
        ]
        date_tos = [
            f.date_to
            for q in queries
            for f in q.email_filters
            if isinstance(f, DateFilter) and f.date_to is not None
        ]
        return (min(date_froms) if date_froms else None, max(date_tos) if date_tos else None)

    def _merge_keywords(
        self, queries: list[SearchQuery]
    ) -> tuple[list[str], bool]:
        """Collect keywords and determine if all are exact match.
        
        Returns:
            Tuple of (keywords list, whether all queries use exact_match).
        """
        keywords = set()
        all_exact_match = True
        for q in queries:
            for f in q.email_filters:
                if isinstance(f, KeywordFilter):
                    keywords.update(f.keywords)
                    if not f.exact_match:
                        all_exact_match = False
        return list(keywords), all_exact_match

    def _apply_query_filters(
        self, emails: list[EmailRecord], query: SearchQuery
    ) -> list[EmailRecord]:
        """Post-filter a list of emails against a single query's keyword and date filters."""
        keyword_filter = next((f for f in query.email_filters if isinstance(f, KeywordFilter)), None)
        date_filter = next((f for f in query.email_filters if isinstance(f, DateFilter)), None)

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

    def _extract_attachment_filters(
        self, attachment_query: AttachmentFilter
    ) -> tuple[list[str], bool]:
        """Extract attachment names and exact_match setting from query filters.
        
        Args:
            attachment_query: AttachmentFilter object with filters.
            
        Returns:
            Tuple of (attachment names list, exact_match boolean).
        """
        attachment_names: list[str] = []
        exact_match = False
        
        for filter_obj in attachment_query.filters:
            if isinstance(filter_obj, KeywordFilter):
                attachment_names.extend(filter_obj.keywords)
                exact_match = filter_obj.exact_match
        
        return attachment_names, exact_match

    def _search_attachment_records(
        self,
        email_records: list[EmailRecord],
        attachment_names: list[str],
        exact_match: bool,
    ) -> list[EmailRecord]:
        """Find emails with matching attachments.
        
        Args:
            email_records: List of EmailRecord objects to search through.
            attachment_names: Attachment names to search for.
            exact_match: Whether to do exact match or substring match.
            
        Returns:
            List of EmailRecord objects with matching attachments, deduplicated.
        """
        if not attachment_names:
            return [r for r in email_records if r.attachments]

        matching_records: list[EmailRecord] = []

        for record in email_records:
            if not record.attachments:
                continue

            for attachment_name in attachment_names:
                for filename in record.attachments:
                    if exact_match:
                        if attachment_name == filename:
                            matching_records.append(record)
                            break
                    else:
                        if attachment_name.lower() in filename.lower():
                            matching_records.append(record)
                            break

                # Break out of attachment_names loop if we found a match
                if matching_records and matching_records[-1] == record:
                    break

        # Remove duplicates while preserving order
        seen = set()
        unique_records = []
        for record in matching_records:
            record_id = id(record)
            if record_id not in seen:
                seen.add(record_id)
                unique_records.append(record)
        
        return unique_records

    def _extract_attachment_contents(
        self,
        email_records: list[EmailRecord],
        attachment_query: AttachmentFilter,
    ) -> dict[Filename, bytes]:
        """Extract attachment content from email records.
        
        Args:
            email_records: List of EmailRecord objects to extract from.
            attachment_query: AttachmentFilter object with filters to apply.
            
        Returns:
            A dict mapping attachment filenames to their binary content.
            
        Raises:
            ValueError: If email record not found in cache or read error occurs.
        """
        # Extract filter parameters from query
        attachment_names, exact_match = self._extract_attachment_filters(attachment_query)
        
        results: dict[Filename, bytes] = {}

        for record in email_records:
            cache_key = self._get_cache_key(record)
            message = self._message_cache.get(cache_key)

            if not message:
                raise EmailNotInCacheError(
                    f"Email record for '{record.subject}' not found in cache. "
                    "Make sure to call filter_emails() first."
                )

            try:
                for i in range(message.Attachments.Count):
                    attachment = message.Attachments.Item(i + 1)
                    filename = attachment.FileName

                    # Skip if searching for specific names and this doesn't match
                    if attachment_names and not self._filename_matches(
                        filename, attachment_names, exact_match
                    ):
                        continue

                    # Save to temp file, read content, then delete temp file
                    with tempfile.TemporaryDirectory(dir=self._temp_dir) as tmpdir:
                        temp_path = Path(tmpdir) / filename
                        attachment.SaveAsFile(str(temp_path))

                        # Read content into memory
                        with open(temp_path, "rb") as f:
                            content = f.read()
                        results[filename] = content

            except EmailNotInCacheError:
                raise
            except Exception as e:
                raise AttachmentReadError(f"Error reading attachments: {e}") from e

        return results

    def _filename_matches(
        self, filename: str, patterns: list[str], exact_match: bool
    ) -> bool:
        """Check if filename matches any pattern.
        
        Args:
            filename: Filename to check.
            patterns: List of patterns to match against.
            exact_match: If True, exact match; if False, substring match.
            
        Returns:
            True if filename matches any pattern.
        """
        for pattern in patterns:
            if exact_match:
                if pattern == filename:
                    return True
            else:
                if pattern.lower() in filename.lower():
                    return True
        return False

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
                    raise FolderNotFoundError(f"Outlook folder '{folder_name}' not found.") from e
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
                raise FolderNotFoundError(f"Outlook folder '{folder_name}' not found.") from e

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

    def _make_email_key(self, record: EmailRecord) -> str:
        """Generate a user-facing key for an email record.

        Args:
            record: EmailRecord to generate key for.

        Returns:
            Key string in 'subject:timestamp' format.
        """
        return f"{record.subject}:{record.received_time.isoformat()}"

    def _extract_attachments_by_email(
        self,
        emails: list[EmailRecord],
        attachment_filter: AttachmentFilter,
    ) -> dict[EmailKey, AttachmentContent]:
        """Extract attachment content from emails, grouped by email key.

        Args:
            emails: List of EmailRecord objects to extract from.
            attachment_filter: AttachmentFilter specifying which attachments to include.

        Returns:
            Dict mapping 'subject:timestamp' keys to {filename: content} dicts.
        """
        attachment_names, exact_match = self._extract_attachment_filters(attachment_filter)
        result: dict[EmailKey, AttachmentContent] = {}

        for record in emails:
            matching = self._search_attachment_records([record], attachment_names, exact_match)
            if not matching:
                continue
            content = self._extract_attachment_contents(matching, attachment_filter)
            if content:
                result[self._make_email_key(record)] = content

        return result
