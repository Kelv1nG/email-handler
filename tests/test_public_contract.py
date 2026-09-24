import inspect

from email_handler.extractors import extract_html_tables, extract_tables_from_msg
from email_handler.protocols import MessageSavingEmailProvider, TableExtractingEmailProvider
from email_handler.protocols.provider import EmailProvider
from email_handler.providers.outlook import OutlookProvider
from email_handler.schemas import ExtractedTable, TableSelector


def test_original_provider_protocol_has_no_new_capabilities():
    members = set(EmailProvider.__dict__)
    assert "extract_tables" not in members
    assert "save_message" not in members


def test_new_capability_signatures_match_outlook_provider():
    assert inspect.signature(
        TableExtractingEmailProvider.extract_tables
    ) == inspect.signature(OutlookProvider.extract_tables)
    assert inspect.signature(
        MessageSavingEmailProvider.save_message
    ) == inspect.signature(OutlookProvider.save_message)


def test_new_public_exports_are_importable():
    assert callable(extract_html_tables)
    assert callable(extract_tables_from_msg)
    assert TableSelector.__name__ == "TableSelector"
    assert ExtractedTable.__name__ == "ExtractedTable"
