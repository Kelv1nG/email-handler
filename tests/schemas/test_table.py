import pytest
from pydantic import ValidationError

from schemas.table import ExtractedTable, TableSelector


def test_selector_defaults_select_all_tables():
    selector = TableSelector()
    assert selector.source_index is None
    assert selector.required_columns == []
    assert selector.occurrence is None
    assert selector.header_row == "auto"


def test_selector_schema_documents_each_parameter():
    properties = TableSelector.model_json_schema()["properties"]
    required_fragments = {
        "source_index": ("zero-based", "DOM order"),
        "required_columns": ("case-insensitively", "whitespace"),
        "occurrence": ("zero-based", "column filtering"),
        "header_row": ("zero-based", '"auto"'),
    }

    for field_name, fragments in required_fragments.items():
        description = properties[field_name].get("description", "").casefold()
        assert all(fragment.casefold() in description for fragment in fragments)


def test_selector_strips_columns_and_rejects_normalized_duplicates():
    assert TableSelector(required_columns=[" Account "]).required_columns == [
        "Account"
    ]
    with pytest.raises(ValidationError, match="unique"):
        TableSelector(required_columns=["Account", " account "])


@pytest.mark.parametrize("columns", [[""], ["   "]])
def test_selector_rejects_blank_required_columns(columns):
    with pytest.raises(ValidationError, match="blank"):
        TableSelector(required_columns=columns)


def test_selector_rejects_two_position_spaces():
    with pytest.raises(ValidationError, match="mutually exclusive"):
        TableSelector(source_index=0, occurrence=0)


@pytest.mark.parametrize(
    ("field", "value"),
    [("source_index", -1), ("occurrence", -1), ("header_row", -1)],
)
def test_selector_rejects_negative_indices(field, value):
    with pytest.raises(ValidationError):
        TableSelector(**{field: value})


def test_extracted_table_requires_rectangular_rows():
    with pytest.raises(ValidationError, match="same width"):
        ExtractedTable(
            source_index=0,
            columns=["A", "B"],
            rows=[["only one"]],
        )
