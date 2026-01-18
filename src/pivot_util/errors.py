"""
Custom error types for the pivot utility.
"""

from __future__ import annotations


class PivotSpecError(Exception):
    """
    Base error for all pivot specification failures.
    """


class ValidationError(PivotSpecError):
    """
    Raised when the pivot specification fails validation checks.
    """


class DestinationError(PivotSpecError):
    """
    Raised when destination sheet handling fails.
    """
