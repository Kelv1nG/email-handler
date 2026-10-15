from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

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


def test_exact_subject_ignores_reply_and_forward_prefixes(monkeypatch):
    class OutlookItems:
        def __init__(self, messages):
            self.messages = messages

        def Sort(self, *_args):
            pass

        def Restrict(self, restriction):
            assert restriction.casefold() == '[subject] = "quarterly report"'
            return self.messages

        def __iter__(self):
            return iter(self.messages)

    def message(subject, sender_email="sender@example.com"):
        return SimpleNamespace(
            Class=43,
            Subject=subject,
            SenderName="Sender",
            SenderEmailAddress=sender_email,
            ReceivedTime=datetime(2026, 9, 23, 10, 0),
            Body="",
            Attachments=SimpleNamespace(Count=0),
        )

    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    items = OutlookItems(
        [
            message("Team Lunch", sender_email=""),
            message("Quarterly Report"),
            message("RE: Quarterly Report"),
            message("RE:Quarterly Report"),
            message("FW: Quarterly Report"),
            message("FW: RE: Quarterly Report"),
            message("FW:RE: Quarterly Report"),
        ]
    )
    monkeypatch.setattr(provider, "_get_folder", lambda _name: SimpleNamespace(Items=items))
    expected_subjects = [
        "Quarterly Report",
        "RE: Quarterly Report",
        "RE:Quarterly Report",
        "FW: Quarterly Report",
        "FW: RE: Quarterly Report",
        "FW:RE: Quarterly Report",
    ]
    for keyword in (
        "quarterly report",
        "RE: Quarterly Report",
        "FW:RE: quarterly report",
    ):
        query = SearchQuery(
            name="reply",
            email_filters=[
                FolderFilter(folder_name="Inbox"),
                KeywordFilter(keywords=[keyword], exact_match=True),
            ],
        )
        records = provider.filter_emails([query])["reply"].records
        assert [record.subject for record in records] == expected_subjects


def test_exact_subject_uses_outlook_equality_before_reading_messages(monkeypatch):
    restrictions = []
    iterated_subjects = []

    class OutlookItems:
        def __init__(self, messages):
            self.messages = messages

        def Sort(self, *_args):
            pass

        def Restrict(self, restriction):
            restrictions.append(restriction)
            assert restriction == '[Subject] = "Quarterly Report"'
            return OutlookItems(
                [
                    message
                    for message in self.messages
                    if message.Subject.removeprefix("RE: ").removeprefix("FW: ").casefold()
                    == "quarterly report"
                ]
            )

        def __iter__(self):
            for message in self.messages:
                iterated_subjects.append(message.Subject)
                yield message

    def message(subject):
        return SimpleNamespace(
            Class=43,
            Subject=subject,
            SenderName="Sender",
            SenderEmailAddress="sender@example.com",
            ReceivedTime=datetime(2026, 9, 23, 10, 0),
            Body="",
            Attachments=SimpleNamespace(Count=0),
        )

    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    items = OutlookItems(
        [
            message("Team Lunch"),
            message("Quarterly Report"),
            message("RE: Quarterly Report"),
            message("FW: Quarterly Report"),
            message("Quarterly Report draft"),
        ]
    )
    monkeypatch.setattr(provider, "_get_folder", lambda _name: SimpleNamespace(Items=items))
    query = SearchQuery(
        name="reports",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["RE: Quarterly Report"], exact_match=True),
        ],
    )

    records = provider.filter_emails([query])["reports"].records

    assert restrictions
    assert "Team Lunch" not in iterated_subjects
    assert [record.subject for record in records] == [
        "Quarterly Report",
        "RE: Quarterly Report",
        "FW: Quarterly Report",
    ]


def test_exact_subject_restriction_escapes_quotes():
    provider = object.__new__(OutlookProvider)
    query = SearchQuery(
        name="quoted",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(
                keywords=['RE: The "Right" Stuff', "Can't Stop"],
                exact_match=True,
            ),
        ],
    )

    assert provider._build_exact_subject_restriction([query]) == (
        '[Subject] = "Can\'t Stop" OR '
        '[Subject] = "The ""Right"" Stuff"'
    )


def test_partial_query_in_same_folder_keeps_other_subjects(monkeypatch):
    class OutlookItems:
        def Sort(self, *_args):
            pass

        def __iter__(self):
            return iter(
                [
                    SimpleNamespace(
                        Class=43,
                        Subject=subject,
                        SenderName="Sender",
                        SenderEmailAddress="sender@example.com",
                        ReceivedTime=datetime(2026, 9, 23, 10, 0),
                        Body="",
                        Attachments=SimpleNamespace(Count=0),
                    )
                    for subject in ("Quarterly Report", "Team Lunch")
                ]
            )

    provider = object.__new__(OutlookProvider)
    provider._message_cache = {}
    monkeypatch.setattr(
        provider, "_get_folder", lambda _name: SimpleNamespace(Items=OutlookItems())
    )
    exact = SearchQuery(
        name="exact",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["Quarterly Report"], exact_match=True),
        ],
    )
    partial = SearchQuery(
        name="partial",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["team"], exact_match=False),
        ],
    )

    results = provider.filter_emails([exact, partial])

    assert [record.subject for record in results["exact"].records] == ["Quarterly Report"]
    assert [record.subject for record in results["partial"].records] == ["Team Lunch"]


def test_substring_subject_search_ignores_only_leading_reply_prefix(monkeypatch):
    provider = object.__new__(OutlookProvider)
    monkeypatch.setattr(
        provider,
        "_search_folder",
        lambda *_args: [
            record("RE: Quarterly Report"),
            record("RE:Quarterly Report"),
            record("Topic RE: Budget"),
        ],
    )
    query = SearchQuery(
        name="reply-text",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["RE:"], exact_match=False),
        ],
    )

    records = provider.filter_emails([query])["reply-text"].records

    assert [record.subject for record in records] == ["Topic RE: Budget"]


def test_substring_subject_search_accepts_prefixed_query(monkeypatch):
    provider = object.__new__(OutlookProvider)
    monkeypatch.setattr(
        provider,
        "_search_folder",
        lambda *_args: [record("Quarterly Report"), record("FW: Quarterly Report")],
    )
    query = SearchQuery(
        name="report",
        email_filters=[
            FolderFilter(folder_name="Inbox"),
            KeywordFilter(keywords=["RE: quarter"], exact_match=False),
        ],
    )

    records = provider.filter_emails([query])["report"].records

    assert [record.subject for record in records] == ["Quarterly Report", "FW: Quarterly Report"]


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
