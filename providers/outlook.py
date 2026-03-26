"""Outlook email provider implementation."""

from datetime import datetime
from pathlib import Path
from typing import Any
import tempfile

import win32com.client

from schemas.email import EmailRecord
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

    def filter_emails(
        self,
        keywords: list[str] | None = None,
        exact_match: bool = False,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
        folder_name: str = "Inbox",
    ) -> list[EmailRecord]:
        """Filter Outlook emails by keywords and/or date range.

        Args:
            keywords: List of keywords to search for in subject and body.
                     If None or empty, no keyword filter is applied.
            exact_match: When True the keyword must match the full subject/body.
                        When False a substring (case-insensitive) match is used.
            date_from: Only include emails received on or after this datetime.
                      Defaults to None (no lower bound).
            date_to: Only include emails received on or before this datetime.
                    Defaults to None (no upper bound).
            folder_name: Outlook folder to search in. Defaults to "Inbox".

        Returns:
            A list of EmailRecord objects matching the filters.

        Raises:
            ValueError: If the folder_name does not exist.
        """
        # Locate the requested folder
        folder = self._get_folder(folder_name)

        # Get messages sorted by received time (newest first)
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)

        # Apply date range filters
        if date_from is not None or date_to is not None:
            messages = self._apply_date_filter(messages, date_from, date_to)

        # Process messages
        results: list[EmailRecord] = []
        for message in messages:
            # Skip non-mail items
            if not self._is_mail_item(message):
                continue

            subject: str = message.Subject or ""

            # Apply keyword filtering
            if keywords and not self._matches_keywords(subject, keywords, exact_match):
                continue

            # Extract attachments
            attachments = self._extract_attachments(message)

            # Create and validate email record
            email_record = EmailRecord(
                subject=subject,
                sender=message.SenderName or "",
                sender_email=message.SenderEmailAddress or "",
                received_time=message.ReceivedTime,
                body=message.Body or "",
                attachments=attachments,
            )

            # Cache the Outlook message object for later attachment retrieval
            cache_key = self._get_cache_key(email_record)
            self._message_cache[cache_key] = message

            results.append(email_record)

        return results

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

    def _apply_date_filter(
        self,
        messages: Any,
        date_from: datetime | None,
        date_to: datetime | None,
    ) -> Any:
        """Build and apply date range restriction.

        Args:
            messages: Outlook messages collection.
            date_from: Start date (inclusive).
            date_to: End date (inclusive).

        Returns:
            Restricted messages collection.
        """
        filters: list[str] = []

        if date_from is not None:
            formatted_date = format_outlook_date(date_from)
            filters.append(f"[ReceivedTime] >= '{formatted_date}'")

        if date_to is not None:
            formatted_date = format_outlook_date(date_to)
            filters.append(f"[ReceivedTime] <= '{formatted_date}'")

        if filters:
            restriction = " AND ".join(filters)
            return messages.Restrict(restriction)

        return messages

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
