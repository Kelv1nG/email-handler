"""Outlook email provider implementation."""

from datetime import datetime
from typing import Any

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
            results.append(email_record)

        return results

    def _get_folder(self, folder_name: str) -> Any:
        """Get folder by name.

        Args:
            folder_name: Name of the folder to retrieve.

        Returns:
            Outlook folder object.

        Raises:
            ValueError: If folder does not exist.
        """
        inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        if folder_name.lower() == "inbox":
            return inbox

        try:
            return inbox.Parent.Folders[folder_name]
        except Exception as e:
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
