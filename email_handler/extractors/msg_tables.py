"""Extract HTML tables from saved Outlook MSG files without Outlook."""

from pathlib import Path

import extract_msg

from email_handler.exceptions import MessageFileReadError
from email_handler.schemas.table import ExtractedTable, TableSelector

from .html_tables import extract_html_tables


def _validated_msg_path(path: str | Path) -> Path:
    candidate = Path(path)
    if not candidate.exists():
        raise FileNotFoundError(candidate)
    if not candidate.is_file():
        raise IsADirectoryError(candidate)
    if candidate.suffix.casefold() != ".msg":
        raise ValueError(f"Expected a .msg file, got '{candidate}'")
    return candidate


def extract_tables_from_msg(
    path: str | Path,
    selector: TableSelector | None = None,
) -> list[ExtractedTable]:
    """Extract selected tables from one saved Outlook MSG file."""
    candidate = Path(path)
    try:
        msg_path = _validated_msg_path(candidate)
    except (FileNotFoundError, IsADirectoryError, ValueError):
        raise
    except OSError as exc:
        raise MessageFileReadError(
            f"Failed to access MSG file '{candidate}': {exc}"
        ) from exc

    try:
        with extract_msg.openMsg(str(msg_path)) as message:
            html_body = message.htmlBody
            if not html_body:
                return []
            return extract_html_tables(html_body, selector)
    except Exception as exc:
        raise MessageFileReadError(
            f"Failed to read tables from MSG file '{msg_path}': {exc}"
        ) from exc
