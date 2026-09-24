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


class DuplicateEmailKeysError(ValueError):
    """Raised when input records would overwrite the same result key."""


class TableExtractionError(RuntimeError):
    """Raised when HTML access or table parsing fails unexpectedly."""


class MessageFileReadError(IOError):
    """Raised when a saved MSG file cannot be decoded or parsed."""


class MessageSaveError(IOError):
    """Raised when Outlook cannot save a cached message as MSG."""
