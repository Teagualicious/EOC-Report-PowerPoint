"""Structured application errors.

``user_message`` carries text that is safe and useful to show in a
messagebox; the exception string itself may contain technical detail
destined for the log file.
"""


class IngestionError(Exception):
    """Base class for expected, user-explainable failures."""

    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message or message


class ParserError(IngestionError):
    """A source file could not be read or contained no usable data."""

