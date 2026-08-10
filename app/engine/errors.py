"""Structured, user-explainable errors for Deck Engine."""

from __future__ import annotations


class IngestionError(Exception):
    """Base class for expected failures safe to show to an analyst."""

    def __init__(self, message: str, user_message: str | None = None,
                 *, code: str | None = None):
        super().__init__(message)
        self.user_message = user_message or message
        self.code = code


class ParserError(IngestionError):
    """A source file could not be read or contained no usable data."""


class UserInputError(IngestionError):
    """The selected dump or staging edit is incomplete or invalid."""


class ConfigError(IngestionError):
    """Settings, mappings, templates, profiles, or rules are inconsistent."""


class EnvironmentError_(IngestionError):
    """The machine or filesystem cannot safely complete the operation."""
