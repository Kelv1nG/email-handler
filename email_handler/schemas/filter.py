from __future__ import annotations
from datetime import datetime
from pydantic import BaseModel, Field

QueryName = str  # Query name identifier


class BaseFilter(BaseModel):
    """Base class for all search filters."""
    ...

class KeywordFilter(BaseFilter):
    """generic class for filtering by keywords"""
    keywords: list[str] = Field(description="Keywords to search for in subject/body")
    exact_match: bool = Field(default=False, description="If True, match full text; if False, substring match")


class EmailFilter(BaseFilter):
    """class for filtering emails"""


class FolderFilter(EmailFilter):
    """Filter emails by folder."""
    folder_name: str = Field(description="Outlook folder to search in")


class DateFilter(EmailFilter):
    """Filter emails by date range."""
    date_from: datetime | None = Field(default=None, description="Search from date (inclusive)")
    date_to: datetime | None = Field(default=None, description="Search to date (inclusive)")


class AttachmentFilter(BaseModel):
    """Filter attachments by filename keywords."""

    filters: list[KeywordFilter] = Field(description="List of keyword filters to match attachment filenames")


class BodyFilter(BaseFilter):
    """Filter emails by body content."""

    keywords: list[str] = Field(description="Keywords to search for in the email body")
    exact_match: bool = Field(
        default=False, description="If True, match full body text; if False, substring match"
    )


class SearchQuery(BaseModel):
    """Composable search query for filtering emails and extracting content."""

    name: QueryName = Field(description="Unique identifier for this query")
    email_filters: list[FolderFilter | KeywordFilter | DateFilter] = Field(
        description="Filters for email search (folder, keywords, date range)"
    )
    attachment_filter: AttachmentFilter | None = Field(
        default=None, description="Optional filter for extracting attachments"
    )
    body_filter: BodyFilter | None = Field(
        default=None, description="Optional filter for email body content"
    )
