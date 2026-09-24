"""Date utility functions for email handling."""

from datetime import datetime


def format_outlook_date(date: datetime) -> str:
    """Format datetime for Outlook restriction strings.

    Args:
        date: Datetime to format.

    Returns:
        Formatted date string for Outlook queries.
    """
    return date.strftime("%m/%d/%Y %H:%M %p")
