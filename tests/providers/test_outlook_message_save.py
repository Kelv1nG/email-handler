from datetime import datetime

import pytest

from email_handler.exceptions import (
    EmailNotInCacheError,
    InvalidSavePathError,
    MessageSaveError,
)
from email_handler.providers.outlook import OutlookProvider
from email_handler.schemas.result import EmailRecord


class FakeMessage:
    def __init__(self, failure=None):
        object.__setattr__(self, "calls", [])
        object.__setattr__(self, "failure", failure)

    def __setattr__(self, name, value):
        raise AssertionError(f"save_message must not assign message.{name}")

    def SaveAs(self, path, save_type):
        self.calls.append((path, save_type))
        if self.failure:
            raise self.failure

    def __getattr__(self, name):
        if name in {"Send", "Move", "Delete", "Save"}:
            raise AssertionError(f"save_message must not call {name}")
        raise AttributeError(name)


def record():
    return EmailRecord(
        subject="Report",
        sender="Sender",
        sender_email="sender@example.com",
        received_time=datetime(2026, 9, 23, 10, 0),
    )


def cached_provider(item, message):
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {provider._get_cache_key(item): message}
    return provider


def test_saves_with_unicode_msg_type_without_mutating_message(tmp_path):
    item = record()
    message = FakeMessage()
    provider = cached_provider(item, message)
    destination = tmp_path / "report.msg"

    assert provider.save_message(item, destination) == destination
    assert message.calls == [(str(destination), 9)]


def test_refuses_existing_destination_without_calling_com(tmp_path):
    item = record()
    message = FakeMessage()
    provider = cached_provider(item, message)
    destination = tmp_path / "report.msg"
    destination.write_bytes(b"keep")

    with pytest.raises(FileExistsError):
        provider.save_message(item, destination)
    assert destination.read_bytes() == b"keep"
    assert message.calls == []


def test_existing_directory_with_msg_suffix_is_also_a_collision(tmp_path):
    item = record()
    message = FakeMessage()
    provider = cached_provider(item, message)
    destination = tmp_path / "report.msg"
    destination.mkdir()

    with pytest.raises(FileExistsError):
        provider.save_message(item, destination)
    assert message.calls == []


@pytest.mark.parametrize("name", ["report.eml", "report"])
def test_requires_msg_suffix(tmp_path, name):
    item = record()
    provider = cached_provider(item, FakeMessage())
    with pytest.raises(InvalidSavePathError, match=r"\.msg"):
        provider.save_message(item, tmp_path / name)


def test_requires_existing_parent(tmp_path):
    item = record()
    provider = cached_provider(item, FakeMessage())
    with pytest.raises(InvalidSavePathError):
        provider.save_message(item, tmp_path / "missing" / "report.msg")


def test_requires_cached_record(tmp_path):
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    with pytest.raises(EmailNotInCacheError):
        provider.save_message(record(), tmp_path / "report.msg")


def test_com_save_failure_is_chained(tmp_path):
    item = record()
    provider = cached_provider(item, FakeMessage(RuntimeError("disk failure")))
    with pytest.raises(MessageSaveError, match="Report") as raised:
        provider.save_message(item, tmp_path / "report.msg")
    assert isinstance(raised.value.__cause__, RuntimeError)
