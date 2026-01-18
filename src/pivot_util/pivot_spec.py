"""
PivotSpec dataclass that validates and generates pivot tables from Excel tables.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

import xlwings as xw

from .constants import DestinationHandling
from .field_specs import ColumnFieldSpec, DataFieldSpec, RowFieldSpec
from .pivot_util import generate_pivot


@dataclass
class PivotSpec:
    """
    Specification for creating a pivot table from an Excel Table (ListObject).

    Args:
        workbook: Existing xlwings workbook object. Caller owns open/save/close.
        table_name: Excel table (ListObject) name.
        pivot_sheet_name: Destination worksheet name.
        pivot_table_name: Pivot table name (must be unique in workbook).
        pivot_top_left_cell: Top-left cell address for the pivot placement.
        row_fields: Row field specs in order (outer to inner).
        column_fields: Column field specs in order (outer to inner).
        data_fields: Data field specs in order.
        destination_handling: How to find/create/clear the destination sheet.
        max_row_fields: Maximum allowed row field depth (default 3).
        max_column_fields: Maximum allowed column field depth (default 3).
        max_data_fields: Maximum allowed data fields (default 20).
        table_style: Optional PivotTable style name (e.g., "PivotStyleMedium9").
        show_row_stripes: Whether to enable row stripes on the PivotTable.

    Raises:
        ValidationError: When the spec fails validation checks.
        DestinationError: When destination handling fails.

    Limitations:
        - Source must be an Excel Table (ListObject), not a raw range.
        - Row/column depth is capped by max_row_fields/max_column_fields.
        - Data field count is capped by max_data_fields.
        - Summary functions are limited to SUM, COUNT, and AVG.
        - If destination_handling is None, an existing non-empty sheet raises an error.
    """

    workbook: xw.Book
    table_name: str
    pivot_sheet_name: str = "Pivot"
    pivot_table_name: str = "PivotTable1"
    pivot_top_left_cell: str = "A3"
    row_fields: List[RowFieldSpec] = field(default_factory=list)
    column_fields: List[ColumnFieldSpec] = field(default_factory=list)
    data_fields: List[DataFieldSpec] = field(default_factory=list)
    destination_handling: Optional[DestinationHandling] = None
    max_row_fields: int = 3
    max_column_fields: int = 3
    max_data_fields: int = 20
    table_style: Optional[str] = None
    show_row_stripes: bool = True

    def generate_pivot(self) -> None:
        """
        Validate the spec and generate the pivot table.

        Returns:
            None

        Raises:
            ValidationError: When the spec fails validation checks.
            DestinationError: When destination handling fails.
        """

        generate_pivot(self)
