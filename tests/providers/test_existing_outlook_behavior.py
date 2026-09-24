from datetime import datetime
from pathlib import Path

from email_handler.providers.outlook import OutlookProvider
from email_handler.schemas.filter import (
    AttachmentFilter,
    BodyFilter,
    FolderFilter,
    KeywordFilter,
    SearchQuery,
)
from email_handler.schemas.result import EmailRecord


def record(subject, body="", attachments=None):
    return EmailRecord(
        subject=subject,
        sender="Sender",
        sender_email="sender@example.com",
        received_time=datetime(2026, 9, 23, 10, 0),
        body=body,
        attachments=attachments or [],
    )


def test_filter_emails_keeps_existing_subject_filtering(monkeypatch):
    provider = object.__new__(OutlookProvider)
    quarterly = record("Quarterly Report")
    unrelated = record("Team Lunch")
    monkeypatch.setattr(
        provider,
        "_search_folder",
        lambda *_args: [quarterly, unrelated],
    )
    query = SearchQuery(
        name="reports",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["quarterly"], exact_match=False),
        ],
    )

    assert provider.filter_emails([query])["reports"].records == [quarterly]


def test_filter_body_keeps_plain_text_body_matching():
    provider = object.__new__(OutlookProvider)
    matching = record("First", body="Contains needle")
    other = record("Second", body="No match")

    result = provider.filter_body(
        [matching, other],
        BodyFilter(keywords=["needle"], exact_match=False),
    )

    assert result == {provider._make_email_key(matching): "Contains needle"}


class FakeAttachment:
    def __init__(self, filename, content):
        self.FileName = filename
        self.content = content

    def SaveAsFile(self, path):
        Path(path).write_bytes(self.content)


class FakeAttachments:
    def __init__(self, *attachments):
        self._attachments = attachments
        self.Count = len(attachments)

    def Item(self, one_based_index):
        return self._attachments[one_based_index - 1]


def test_filter_attachments_keeps_filename_filtering_and_binary_reads(tmp_path):
    item = record(
        "Report",
        attachments=["Monthly Report.CSV", "ignore.txt"],
    )
    message = type(
        "Message",
        (),
        {
            "Attachments": FakeAttachments(
                FakeAttachment("Monthly Report.CSV", b"csv-content"),
                FakeAttachment("ignore.txt", b"ignore"),
            )
        },
    )()
    provider = object.__new__(OutlookProvider)
    provider._message_cache = {provider._get_cache_key(item): message}
    provider._temp_dir = str(tmp_path)
    attachment_filter = AttachmentFilter(
        filters=[KeywordFilter(keywords=["report"], exact_match=False)]
    )

    assert provider.filter_attachments([item], attachment_filter) == {
        "Monthly Report.CSV": b"csv-content"
    }
