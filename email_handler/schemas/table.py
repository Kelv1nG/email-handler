"""Validated selectors and results for HTML table extraction."""

from typing import Literal, Self

from pydantic import BaseModel, Field, field_validator, model_validator


def _comparison_name(value: str) -> str:
    return " ".join(value.split()).casefold()


class TableSelector(BaseModel):
    """Select HTML tables by position, columns, and header interpretation.

    ``source_index`` selects from every HTML table before column filtering,
    while ``occurrence`` selects from the tables left after column filtering.
    The two position spaces are mutually exclusive. Omitting both returns all
    tables that satisfy ``required_columns``.

    Args:
        source_index: Absolute zero-based table position in HTML DOM order.
            Empty or non-data tables still occupy a position. Cannot be used
            together with ``occurrence``.
        required_columns: Column names that every returned table must contain.
            Names are matched case-insensitively after whitespace normalization.
            An empty list accepts any set of columns.
        occurrence: Zero-based position among tables remaining after column
            filtering. Cannot be used together with ``source_index``.
        header_row: Zero-based logical row to use as the column header, or
            ``"auto"`` to infer it from ``thead``, header cells, or the first
            non-empty row.
    """

    source_index: int | None = Field(
        default=None,
        ge=0,
        description="Absolute zero-based table position in HTML DOM order.",
    )
    required_columns: list[str] = Field(
        default_factory=list,
        description=(
            "Column names matched case-insensitively after whitespace "
            "normalization."
        ),
    )
    occurrence: int | None = Field(
        default=None,
        ge=0,
        description=(
            "Zero-based position among matching tables after column filtering."
        ),
    )
    header_row: int | Literal["auto"] = Field(
        default="auto",
        description=(
            'Zero-based logical header row, or "auto" for automatic inference.'
        ),
    )

    @field_validator("required_columns")
    @classmethod
    def validate_required_columns(cls, values: list[str]) -> list[str]:
        stripped = [value.strip() for value in values]
        if any(not value for value in stripped):
            raise ValueError("required columns cannot contain blank names")
        normalized = [_comparison_name(value) for value in stripped]
        if len(normalized) != len(set(normalized)):
            raise ValueError("required columns must be unique after normalization")
        return stripped

    @field_validator("header_row")
    @classmethod
    def validate_header_row(
        cls, value: int | Literal["auto"]
    ) -> int | Literal["auto"]:
        if isinstance(value, int) and value < 0:
            raise ValueError("header_row must be non-negative")
        return value

    @model_validator(mode="after")
    def validate_position_space(self) -> Self:
        if self.source_index is not None and self.occurrence is not None:
            raise ValueError("source_index and occurrence are mutually exclusive")
        return self


class ExtractedTable(BaseModel):
    """A rectangular, source-ordered table extracted from an HTML body."""

    source_index: int = Field(ge=0)
    caption: str | None = None
    columns: list[str]
    rows: list[list[str | None]]

    @model_validator(mode="after")
    def validate_rectangular_rows(self) -> Self:
        width = len(self.columns)
        if any(len(row) != width for row in self.rows):
            raise ValueError("all rows must have the same width as columns")
        return self
