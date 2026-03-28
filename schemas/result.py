"""Result schemas for email queries."""

from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Literal
from pydantic import BaseModel, Field, field_validator

from schemas.filter import QueryName, SearchQuery

EmailKey = str  # "subject:timestamp" format
AttachmentContent = dict[str, bytes]  # {filename: bytes}
BodyContent = str  # Email body text
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

    name: QueryName = Field(description="Query name")
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

    def save_attachments(
        self,
        path_mapper: Callable[["EmailRecord", Filename], Path | None],
        on_collision: Literal["overwrite", "skip", "error"] = "error",
    ) -> dict[Path, Filename]:
        """Save in-memory attachments to disk using a path mapper.

        Args:
            path_mapper: Called for each (EmailRecord, filename) pair. Return the
                destination Path to save the file, or None to skip it.
            on_collision: What to do when two files resolve to the same destination.
                "overwrite" silently overwrites, "skip" keeps the first, "error" raises.

        Returns:
            Dict mapping saved destination Path to the original filename.
        """
        email_by_key: dict[EmailKey, EmailRecord] = {
            f"{r.subject}:{r.received_time.isoformat()}": r for r in self.emails
        }
        saved: dict[Path, Filename] = {}
        for email_key, files in self.attachments.items():
            email = email_by_key[email_key]
            for filename, content in files.items():
                dest = path_mapper(email, filename)
                if dest is None:
                    continue
                if dest in saved:
                    if on_collision == "error":
                        raise FileExistsError(
                            f"Collision: '{filename}' and '{saved[dest]}' both resolve to '{dest}'"
                        )
                    elif on_collision == "skip":
                        continue
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_bytes(content)
                saved[dest] = filename
        return saved

