"""Validated selectors and results for HTML table extraction."""

from typing import Literal, Self

from pydantic import BaseModel, Field, field_validator, model_validator


def _comparison_name(value: str) -> str:
    return " ".join(value.split()).casefold()


class TableSelector(BaseModel):
    """Select tables by absolute position, columns, or filtered occurrence."""

    source_index: int | None = Field(default=None, ge=0)
    required_columns: list[str] = Field(default_factory=list)
    occurrence: int | None = Field(default=None, ge=0)
    header_row: int | Literal["auto"] = "auto"

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
