"""Result schemas for email queries."""

from pydantic import BaseModel, Field

from schemas.email import EmailRecord
from schemas.filter import SearchQuery


class QueryResult(BaseModel):
    """Results grouped by query."""

    query: SearchQuery
    records: list[EmailRecord] = Field(description="Email records matching this query")


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


