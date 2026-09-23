"""Optional protocol for providers that persist complete messages."""

from pathlib import Path
from typing import Protocol

from schemas.result import EmailRecord


class MessageSavingEmailProvider(Protocol):
    def save_message(
        self,
        email_record: EmailRecord,
        save_path: str | Path,
    ) -> Path:
        """Save one complete message and return its destination."""
