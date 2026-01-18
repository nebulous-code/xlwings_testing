"""
Custom error types for the pivot utility.
"""

from __future__ import annotations


class PivotBuilderError(Exception):
    """
    Base error for all pivot builder failures.
    """


class ValidationError(PivotBuilderError):
    """
    Raised when the pivot specification fails validation checks.
    """


class DestinationError(PivotBuilderError):
    """
    Raised when destination sheet handling fails.
    """
