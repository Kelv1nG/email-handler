"""Result schemas for email queries."""

from datetime import datetime
from pydantic import BaseModel, Field, field_validator

from schemas.filter import SearchQuery


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

