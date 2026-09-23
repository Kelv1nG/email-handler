"""Provider-independent email content extractors."""

from .html_tables import extract_html_tables

__all__ = ["extract_html_tables"]
