from email_handler.schemas.result import ExtractionResult


def test_extraction_result_print_summarizes_attachment_bytes(capsys):
    content = b"\x00\xff\x01\x02" * 4
    result = ExtractionResult(
        name="report",
        attachments={"mail-key": {"report.bin": content}},
    )

    representation = repr(result)
    print(result)
    printed = capsys.readouterr().out

    for display in (representation, printed):
        assert "mail-key" in display
        assert "report.bin" in display
        assert "16 bytes" in display
        assert "\\x00" not in display
        assert "b'" not in display

    assert result.model_dump()["attachments"]["mail-key"]["report.bin"] == content
