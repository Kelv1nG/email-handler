"""Protocol definitions for email providers."""

from .provider import EmailProvider
from .table_provider import TableExtractingEmailProvider

__all__ = ["EmailProvider", "TableExtractingEmailProvider"]
