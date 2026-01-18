"""
Core pivot creation logic that reads a PivotSpec and builds the pivot table.
"""

from __future__ import annotations

from typing import Iterable, Optional, TYPE_CHECKING

import xlwings as xw

from .constants import (
    DestinationHandling,
    SummaryFunction,
    XL_COLUMN_FIELD,
    XL_DATA_FIELD,
    XL_DATABASE,
    XL_ROW_FIELD,
    XL_SUM,
    XL_COUNT,
    XL_AVG,
)
from .errors import DestinationError, ValidationError
from .field_specs import ColumnFieldSpec, DataFieldSpec, RowFieldSpec
if TYPE_CHECKING:
    from .pivot_spec import PivotSpec


def generate_pivot(spec: PivotSpec) -> None:
    """
    Validate and generate a pivot table based on the provided spec.

    Args:
        spec: PivotSpec with field definitions and destination preferences.

    Returns:
        None

    Raises:
        ValidationError: When the spec fails validation checks.
        DestinationError: When destination handling fails.
    """

    list_object = _validate_and_get_table(spec)
    pivot_sheet = _resolve_destination_sheet(spec)

    # Optional clearing behavior based on destination handling.
    if spec.destination_handling == DestinationHandling.EXISTING_FORCE_CLEAR:
        pivot_sheet.clear()

    # The pivot cache source range should be the Table's range (includes headers).
    source_range = list_object.Range  # COM Range

    # Build the pivot cache and pivot table at the requested top-left cell.
    pivot_cache = spec.workbook.api.PivotCaches().Create(XL_DATABASE, source_range)  # type: ignore[attr-defined]
    dest = pivot_sheet.range(spec.pivot_top_left_cell).api  # COM Range
    pivot_table = pivot_cache.CreatePivotTable(dest, spec.pivot_table_name)

    # Configure row fields (outer to inner).
    for idx, row_field in enumerate(spec.row_fields, start=1):
        pf = pivot_table.PivotFields(row_field.name)
        pf.Orientation = XL_ROW_FIELD
        pf.Position = idx
        if row_field.caption:
            pf.Caption = row_field.caption

    # Configure column fields (outer to inner).
    for idx, col_field in enumerate(spec.column_fields, start=1):
        pf = pivot_table.PivotFields(col_field.name)
        pf.Orientation = XL_COLUMN_FIELD
        pf.Position = idx
        if col_field.caption:
            pf.Caption = col_field.caption

    # Configure data fields with summary functions and formatting.
    for data_field in spec.data_fields:
        pf = pivot_table.PivotFields(data_field.name)
        pf.Orientation = XL_DATA_FIELD
        pf.Function = _summary_function_to_excel(data_field.function)
        if data_field.caption:
            pf.Caption = data_field.caption
        if data_field.number_format:
            pf.NumberFormat = data_field.number_format

    if spec.table_style:
        pivot_table.TableStyle2 = spec.table_style
    pivot_table.ShowTableStyleRowStripes = bool(spec.show_row_stripes)


def _validate_and_get_table(spec: PivotSpec):
    _validate_spec_inputs(spec)
    list_object = _find_list_object(spec.workbook, spec.table_name)
    if list_object is None:
        raise ValidationError(f"Table '{spec.table_name}' not found in workbook.")

    column_names = _list_object_column_names(list_object)
    _validate_unique_column_names(column_names)
    _validate_field_names_exist(spec.row_fields, spec.column_fields, spec.data_fields, column_names)
    _validate_pivot_table_name_unique(spec.workbook, spec.pivot_table_name)

    return list_object


def _validate_spec_inputs(spec: PivotSpec) -> None:
    if len(spec.row_fields) > spec.max_row_fields:
        raise ValidationError(
            f"Row field depth {len(spec.row_fields)} exceeds max of {spec.max_row_fields}."
        )
    if len(spec.column_fields) > spec.max_column_fields:
        raise ValidationError(
            f"Column field depth {len(spec.column_fields)} exceeds max of {spec.max_column_fields}."
        )
    if len(spec.data_fields) == 0:
        raise ValidationError("At least one data field is required.")
    if len(spec.data_fields) > spec.max_data_fields:
        raise ValidationError(
            f"Data fields count {len(spec.data_fields)} exceeds max of {spec.max_data_fields}."
        )
    for data_field in spec.data_fields:
        if not isinstance(data_field.function, SummaryFunction):
            raise ValidationError(
                f"Data field '{data_field.name}' uses unsupported function '{data_field.function}'."
            )


def _resolve_destination_sheet(spec: PivotSpec) -> xw.Sheet:
    handling = spec.destination_handling

    try:
        existing_sheet = spec.workbook.sheets[spec.pivot_sheet_name]
    except Exception:
        existing_sheet = None

    if existing_sheet is not None:
        if handling == DestinationHandling.NEW:
            raise DestinationError(f"Sheet '{spec.pivot_sheet_name}' already exists.")
        if handling == DestinationHandling.EXISTING_NO_CLEAR:
            return existing_sheet
        if handling == DestinationHandling.EXISTING_FORCE_CLEAR:
            return existing_sheet
        if handling in (DestinationHandling.EXISTING_CLEAR, None):
            if _is_sheet_empty(existing_sheet):
                return existing_sheet
            raise DestinationError(
                f"Sheet '{spec.pivot_sheet_name}' is not empty under ExistingClear rules."
            )
    else:
        if handling in (
            DestinationHandling.EXISTING_CLEAR,
            DestinationHandling.EXISTING_FORCE_CLEAR,
            DestinationHandling.EXISTING_NO_CLEAR,
        ):
            raise DestinationError(f"Sheet '{spec.pivot_sheet_name}' was not found.")

    # handling is None or NEW and the sheet does not exist.
    return spec.workbook.sheets.add(spec.pivot_sheet_name, after=spec.workbook.sheets[-1])


def _find_list_object(workbook: xw.Book, table_name: str):
    for sheet in workbook.sheets:
        try:
            return sheet.api.ListObjects(table_name)  # type: ignore[attr-defined]
        except Exception:
            continue
    return None


def _list_object_column_names(list_object) -> list[str]:
    try:
        columns = list_object.ListColumns
        return [columns.Item(i).Name for i in range(1, columns.Count + 1)]
    except Exception:
        return []


def _validate_unique_column_names(column_names: Iterable[str]) -> None:
    lowered = [name.strip().lower() for name in column_names]
    if len(lowered) != len(set(lowered)):
        raise ValidationError("Source table has duplicate column names.")


def _validate_field_names_exist(
    row_fields: Iterable[RowFieldSpec],
    column_fields: Iterable[ColumnFieldSpec],
    data_fields: Iterable[DataFieldSpec],
    column_names: Iterable[str],
) -> None:
    existing = {name.strip().lower() for name in column_names}
    requested = []
    requested.extend([rf.name for rf in row_fields])
    requested.extend([cf.name for cf in column_fields])
    requested.extend([df.name for df in data_fields])
    missing = [name for name in requested if name.strip().lower() not in existing]
    if missing:
        raise ValidationError(f"Fields not found in table: {', '.join(missing)}")


def _validate_pivot_table_name_unique(workbook: xw.Book, pivot_table_name: str) -> None:
    for sheet in workbook.sheets:
        try:
            pivot_tables = sheet.api.PivotTables()
            count = pivot_tables.Count
        except Exception:
            continue
        for i in range(1, count + 1):
            if pivot_tables.Item(i).Name == pivot_table_name:
                raise ValidationError(
                    f"Pivot table name '{pivot_table_name}' already exists."
                )


def _summary_function_to_excel(function: SummaryFunction) -> int:
    if function == SummaryFunction.SUM:
        return XL_SUM
    if function == SummaryFunction.COUNT:
        return XL_COUNT
    if function == SummaryFunction.AVG:
        return XL_AVG
    raise ValidationError(f"Unsupported summary function: {function}")


def _is_sheet_empty(sheet: xw.Sheet) -> bool:
    values = sheet.used_range.value
    return _is_empty_value(values)


def _is_empty_value(value: Optional[object]) -> bool:
    if value is None:
        return True
    if isinstance(value, (list, tuple)):
        return all(_is_empty_value(item) for item in value)
    if isinstance(value, str):
        return value.strip() == ""
    return False
