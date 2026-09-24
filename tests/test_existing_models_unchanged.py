from datetime import datetime

from email_handler.schemas.filter import SearchQuery
from email_handler.schemas.result import EmailRecord, ExtractionResult, QueryResult


def test_existing_pydantic_field_sets_are_unchanged():
    assert set(EmailRecord.model_fields) == {
        "subject",
        "sender",
        "sender_email",
        "received_time",
        "body",
        "attachments",
    }
    assert set(SearchQuery.model_fields) == {
        "name",
        "email_filters",
        "attachment_filter",
        "body_filter",
    }
    assert set(QueryResult.model_fields) == {"query", "records"}
    assert set(ExtractionResult.model_fields) == {
        "name",
        "emails",
        "attachments",
        "bodies",
        "error",
    }


def test_existing_models_keep_exact_serialized_shapes_and_defaults():
    received = datetime(2026, 9, 23, 10, 0)
    email = EmailRecord(
        sender_email="sender@example.com",
        received_time=received,
    )
    assert email.model_dump() == {
        "subject": "",
        "sender": "",
        "sender_email": "sender@example.com",
        "received_time": received,
        "body": "",
        "attachments": [],
    }

    query = SearchQuery(name="daily", email_filters=[])
    assert query.model_dump() == {
        "name": "daily",
        "email_filters": [],
        "attachment_filter": None,
        "body_filter": None,
    }

    result = ExtractionResult(
        name="daily",
        emails=[email],
        attachments={"report-key": {"report.csv": b"a,b\n1,2\n"}},
        bodies={"report-key": "plain body"},
    )
    assert result.model_dump() == {
        "name": "daily",
        "emails": [email.model_dump()],
        "attachments": {"report-key": {"report.csv": b"a,b\n1,2\n"}},
        "bodies": {"report-key": "plain body"},
        "error": None,
    }
