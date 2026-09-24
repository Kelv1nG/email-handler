"""Protocol definitions for email providers."""

from .provider import EmailProvider
from .message_saver import MessageSavingEmailProvider
from .table_provider import TableExtractingEmailProvider

__all__ = [
    "EmailProvider",
    "MessageSavingEmailProvider",
    "TableExtractingEmailProvider",
]
