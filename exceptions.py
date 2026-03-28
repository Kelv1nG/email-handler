"""Custom exceptions for the email-extractor service layer."""


class DuplicateQueryNamesError(ValueError):
    """Raised when two or more SearchQuery objects share the same name."""


class FolderNotFoundError(LookupError):
    """Raised when a referenced Outlook folder does not exist."""


class InvalidSavePathError(ValueError):
    """Raised when the provided save path is missing or not a directory."""


class EmailNotInCacheError(LookupError):
    """Raised when an EmailRecord has no corresponding cached Outlook message.
    
    This typically means filter_emails() was not called before attempting
    to read attachments or body content.
    """


class AttachmentReadError(IOError):
    """Raised when an attachment cannot be read from an Outlook message."""
