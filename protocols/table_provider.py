"""Optional protocol for providers that expose structured table extraction."""

from typing import Protocol

from schemas.result import EmailKey, EmailRecord
from schemas.table import ExtractedTable, TableSelector


class TableExtractingEmailProvider(Protocol):
    def extract_tables(
        self,
        email_records: list[EmailRecord],
        selector: TableSelector | None = None,
    ) -> dict[EmailKey, list[ExtractedTable]]:
        """Extract selected tables from the supplied email records."""
