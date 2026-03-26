from __future__ import annotations
from datetime import datetime
from pydantic import BaseModel, Field


class BaseFilter(BaseModel):
    """Base class for all search filters."""
    ...


class FolderFilter(BaseFilter):
    """Filter emails by folder."""
    folder_name: str = Field(description="Outlook folder to search in")


class KeywordFilter(BaseFilter):
    """Filter emails by keywords in subject/body."""
    keywords: list[str] = Field(description="Keywords to search for in subject/body")
    exact_match: bool = Field(default=False, description="If True, match full text; if False, substring match")


class DateFilter(BaseFilter):
    """Filter emails by date range."""
    date_from: datetime | None = Field(default=None, description="Search from date (inclusive)")
    date_to: datetime | None = Field(default=None, description="Search to date (inclusive)")


class SearchQuery(BaseModel):
    """Composeable search query with multiple filter types."""
    filters: list[BaseFilter] = Field(description="List of filters to apply (FolderFilter, KeywordFilter, DateFilter, etc.)")

