"""Result schemas for email queries."""

from datetime import datetime
from pydantic import BaseModel, Field, field_validator

from schemas.filter import SearchQuery

EmailKey = str  # "subject:timestamp" format
AttachmentContent = dict[str, bytes]  # {filename: bytes}
BodyContent = str  # Email body text
QueryName = str  # Query name identifier
Filename = str  # Attachment filename


class EmailRecord(BaseModel):
    """Validated email record from email providers."""

    subject: str = Field(default="", description="Email subject line")
    sender: str = Field(default="", description="Sender display name")
    sender_email: str = Field(description="Sender email address")
    received_time: datetime = Field(description="Time email was received")
    body: str = Field(default="", description="Email body text")
    attachments: list[str] = Field(
        default_factory=list, description="List of attachment filenames"
    )

    @field_validator("sender_email", mode="before")
    @classmethod
    def validate_sender_email(cls, v):
        """Ensure sender_email is not empty."""
        if not v or not str(v).strip():
            raise ValueError("sender_email cannot be empty")
        return str(v).strip()


class QueryResult(BaseModel):
    """Results grouped by query."""

    query: SearchQuery
    records: list[EmailRecord] = Field(description="Email records matching this query")


class ExtractionResult(BaseModel):
    """Result from a full extraction pipeline (filter emails + attachments + body)."""

    emails: list[EmailRecord] = Field(
        default_factory=list, description="Emails matching the query"
    )
    attachments: dict[EmailKey, AttachmentContent] = Field(
        default_factory=dict,
        description="Attachments keyed by email key, with filename to raw bytes mapping",
    )
    bodies: dict[EmailKey, BodyContent] = Field(
        default_factory=dict,
        description="Body text keyed by email key for emails matching the body filter",
    )
    error: str | None = Field(default=None, description="Error message if extraction failed")

