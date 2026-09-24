from datetime import datetime

import pytest

from email_handler.exceptions import (
    DuplicateEmailKeysError,
    EmailNotInCacheError,
    TableExtractionError,
)
from email_handler.providers.outlook import OutlookProvider
from email_handler.schemas.result import EmailRecord
from email_handler.schemas.table import TableSelector


def record(sender="sender@example.com", subject="Report", minute=0):
    return EmailRecord(
        subject=subject,
        sender="Sender",
        sender_email=sender,
        received_time=datetime(2026, 9, 23, 10, minute),
        body="plain text remains unchanged",
    )


def provider_with_messages(records_and_messages):
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {
        provider._get_cache_key(item): message
        for item, message in records_and_messages
    }
    return provider


def test_extracts_live_html_and_preserves_plain_body():
    item = record()
    message = type(
        "Message",
        (),
        {"HTMLBody": "<table><tr><th>A</th></tr><tr><td>1</td></tr></table>"},
    )()
    provider = provider_with_messages([(item, message)])

    result = provider.extract_tables(
        [item], TableSelector(required_columns=["A"])
    )

    assert result[provider._make_email_key(item)][0].rows == [["1"]]
    assert item.body == "plain text remains unchanged"


@pytest.mark.parametrize("html_body", [None, ""])
def test_live_empty_html_keeps_email_key(html_body):
    item = record()
    provider = provider_with_messages(
        [(item, type("Message", (), {"HTMLBody": html_body})())]
    )
    assert provider.extract_tables([item]) == {provider._make_email_key(item): []}


@pytest.mark.parametrize("html_body", [b"<table></table>", 123])
def test_live_html_body_requires_com_string(html_body):
    item = record()
    message = type("Message", (), {"HTMLBody": html_body})()
    provider = provider_with_messages([(item, message)])

    with pytest.raises(TableExtractionError, match="Report") as raised:
        provider.extract_tables([item])

    assert isinstance(raised.value.__cause__, TypeError)


def test_missing_cached_message_is_explicit():
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    with pytest.raises(EmailNotInCacheError):
        provider.extract_tables([record()])


def test_duplicate_user_facing_keys_fail_before_overwrite():
    first = record("first@example.com")
    second = record("second@example.com")
    provider = provider_with_messages(
        [
            (first, type("Message", (), {"HTMLBody": ""})()),
            (second, type("Message", (), {"HTMLBody": ""})()),
        ]
    )
    with pytest.raises(DuplicateEmailKeysError):
        provider.extract_tables([first, second])


def test_html_property_failure_is_chained():
    class BrokenMessage:
        @property
        def HTMLBody(self):
            raise RuntimeError("COM failure")

    item = record()
    provider = provider_with_messages([(item, BrokenMessage())])
    with pytest.raises(TableExtractionError, match="Report") as raised:
        provider.extract_tables([item])
    assert isinstance(raised.value.__cause__, RuntimeError)


def test_shared_parser_failure_gets_email_context_and_is_chained(monkeypatch):
    item = record()
    message = type("Message", (), {"HTMLBody": "<table></table>"})()
    provider = provider_with_messages([(item, message)])
    parser_error = TableExtractionError("parser exploded")

    def fail(_html, _selector):
        raise parser_error

    monkeypatch.setattr("email_handler.providers.outlook.extract_html_tables", fail)
    with pytest.raises(TableExtractionError, match="Report") as raised:
        provider.extract_tables([item])

    assert raised.value.__cause__ is parser_error


def test_preserves_email_order_and_table_source_order():
    first = record(subject="First")
    second = record(subject="Second")
    first_message = type(
        "Message",
        (),
        {
            "HTMLBody": (
                "<table><tr><th>A</th></tr></table>"
                "<table><tr><th>B</th></tr></table>"
            )
        },
    )()
    second_message = type(
        "Message",
        (),
        {"HTMLBody": "<table><tr><th>C</th></tr></table>"},
    )()
    provider = provider_with_messages(
        [(first, first_message), (second, second_message)]
    )

    result = provider.extract_tables([first, second])

    assert list(result) == [
        provider._make_email_key(first),
        provider._make_email_key(second),
    ]
    assert [
        table.source_index
        for table in result[provider._make_email_key(first)]
    ] == [0, 1]


def test_first_failure_does_not_touch_later_message():
    class BrokenFirst:
        @property
        def HTMLBody(self):
            raise RuntimeError("first failed")

    class LaterTripwire:
        touched = False

        @property
        def HTMLBody(self):
            self.touched = True
            raise AssertionError("later message must not be touched")

    first = record(subject="First")
    second = record(subject="Second")
    later = LaterTripwire()
    provider = provider_with_messages([(first, BrokenFirst()), (second, later)])

    with pytest.raises(TableExtractionError, match="First"):
        provider.extract_tables([first, second])

    assert later.touched is False
