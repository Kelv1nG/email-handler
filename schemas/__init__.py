"""Schema definitions for email data."""

from .email import EmailRecord
from .filter import BaseFilter, FolderFilter, KeywordFilter, DateFilter, SearchQuery
from .result import QueryResult

__all__ = [
    "EmailRecord",
    "BaseFilter",
    "FolderFilter",
    "KeywordFilter",
    "DateFilter",
    "SearchQuery",
    "QueryResult",
]
