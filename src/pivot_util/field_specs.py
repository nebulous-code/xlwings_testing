"""
Field specification objects for building pivot tables.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .constants import SummaryFunction


@dataclass(frozen=True)
class RowFieldSpec:
    """
    Row field definition for a pivot table.

    Args:
        name: Source column name from the Excel table.
        caption: Optional display name for the field in the pivot.
    """

    name: str
    caption: Optional[str] = None


@dataclass(frozen=True)
class ColumnFieldSpec:
    """
    Column field definition for a pivot table.

    Args:
        name: Source column name from the Excel table.
        caption: Optional display name for the field in the pivot.
    """

    name: str
    caption: Optional[str] = None


@dataclass(frozen=True)
class DataFieldSpec:
    """
    Data field definition for a pivot table.

    Args:
        name: Source column name from the Excel table.
        function: Summary function for aggregation.
        caption: Optional display name for the data field in the pivot.
        number_format: Optional Excel number format string.
    """

    name: str
    function: SummaryFunction
    caption: Optional[str] = None
    number_format: Optional[str] = None
