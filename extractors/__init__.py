"""Provider-independent email content extractors."""

from .html_tables import extract_html_tables
from .msg_tables import extract_tables_from_msg

__all__ = ["extract_html_tables", "extract_tables_from_msg"]
