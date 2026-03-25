"""Custom validators for email records."""

import re
from typing import Any


def validate_email_format(email: str) -> bool:
    """Validate email address format.

    Args:
        email: Email address to validate.

    Returns:
        True if valid, False otherwise.
    """
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None


def validate_keywords(keywords: list[str] | None) -> bool:
    """Validate keyword list.

    Args:
        keywords: List of keywords to validate.

    Returns:
        True if valid or None, False otherwise.
    """
    if keywords is None:
        return True
    if not isinstance(keywords, list):
        return False
    return all(isinstance(kw, str) and kw.strip() for kw in keywords)
