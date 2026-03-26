"""Schema definitions for email data."""

from .filter import BaseFilter, FolderFilter, KeywordFilter, DateFilter, SearchQuery
from .result import EmailRecord, QueryResult

__all__ = [
    "EmailRecord",
    "BaseFilter",
    "FolderFilter",
    "KeywordFilter",
    "DateFilter",
    "SearchQuery",
    "QueryResult",
]
