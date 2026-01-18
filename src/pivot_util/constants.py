"""
Excel/COM constants and enums used by the pivot utility.
"""

from __future__ import annotations

from enum import Enum


# Excel PivotField orientation constants (COM)
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_DATA_FIELD = 4

# Excel summary function constants (COM)
XL_SUM = -4157
XL_COUNT = -4112
XL_AVG = -4106

# Excel PivotCache source type (COM)
XL_DATABASE = 1


class DestinationHandling(Enum):
    """
    Defines how the destination sheet should be found or created.
    """

    EXISTING_CLEAR = "existing_clear"
    EXISTING_FORCE_CLEAR = "existing_force_clear"
    EXISTING_NO_CLEAR = "existing_no_clear"
    NEW = "new"


class SummaryFunction(Enum):
    """
    Allowed summary functions for data fields.
    """

    SUM = "sum"
    COUNT = "count"
    AVG = "avg"
