"""
Unit tests for pivot_util with mocked workbook and COM objects.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

import pytest

from pivot_util.constants import DestinationHandling, SummaryFunction, XL_AVG, XL_COUNT, XL_SUM
from pivot_util.errors import DestinationError, ValidationError
from pivot_util.fields import ColumnField, DataField, RowField
from pivot_util.pivot_builder import PivotBuilder
from pivot_util.pivot_util import (
    _find_list_object,
    _is_empty_value,
    _list_object_column_names,
    _resolve_destination_sheet,
    _summary_function_to_excel,
    _validate_and_get_table,
    _validate_field_names_exist,
    _validate_pivot_table_name_unique,
    _validate_spec_inputs,
    _validate_unique_column_names,
    generate_pivot,
)


class FakeRange:
    def __init__(self, api_value: str = "dest") -> None:
        self.api = api_value


class FakeListColumn:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeListColumns:
    def __init__(self, names: List[str]) -> None:
        self._names = names
        self.Count = len(names)

    def Item(self, i: int) -> FakeListColumn:
        return FakeListColumn(self._names[i - 1])


class FakeListObject:
    def __init__(self, column_names: List[str]) -> None:
        self.Range = FakeRange("source")
        self.ListColumns = FakeListColumns(column_names)


class FakeListObjectBroken:
    @property
    def ListColumns(self):
        raise RuntimeError("broken")


class FakePivotField:
    def __init__(self) -> None:
        self.Orientation: Optional[int] = None
        self.Position: Optional[int] = None
        self.Function: Optional[int] = None
        self.Caption: Optional[str] = None
        self.NumberFormat: Optional[str] = None


class FakePivotTable:
    def __init__(self) -> None:
        self._fields: Dict[str, FakePivotField] = {}
        self.TableStyle2: Optional[str] = None
        self.ShowTableStyleRowStripes: Optional[bool] = None

    def PivotFields(self, name: str) -> FakePivotField:
        if name not in self._fields:
            self._fields[name] = FakePivotField()
        return self._fields[name]


class FakePivotCache:
    def __init__(self) -> None:
        self.created_with: Optional[tuple] = None
        self.table: Optional[FakePivotTable] = None

    def CreatePivotTable(self, dest, name: str) -> FakePivotTable:
        self.created_with = (dest, name)
        self.table = FakePivotTable()
        return self.table


class FakePivotCaches:
    def __init__(self) -> None:
        self.last_create_args: Optional[tuple] = None
        self.cache = FakePivotCache()

    def Create(self, source_type: int, source_range) -> FakePivotCache:
        self.last_create_args = (source_type, source_range)
        return self.cache


class FakePivotTables:
    def __init__(self, names: List[str]) -> None:
        self._names = names
        self.Count = len(names)

    def Item(self, i: int):
        return type("Pivot", (), {"Name": self._names[i - 1]})


class FakeSheetApi:
    def __init__(self, list_objects: Dict[str, FakeListObject], pivot_table_names: List[str]):
        self._list_objects = list_objects
        self._pivot_table_names = pivot_table_names

    def ListObjects(self, name: str):
        if name not in self._list_objects:
            raise Exception("not found")
        return self._list_objects[name]

    def PivotTables(self):
        return FakePivotTables(self._pivot_table_names)


class FakeUsedRange:
    def __init__(self, value) -> None:
        self.value = value


class FakeSheet:
    def __init__(
        self,
        name: str,
        list_objects: Optional[Dict[str, FakeListObject]] = None,
        pivot_table_names: Optional[List[str]] = None,
        used_range_value=None,
    ) -> None:
        self.name = name
        self._cleared = False
        self.used_range = FakeUsedRange(used_range_value)
        self.api = FakeSheetApi(list_objects or {}, pivot_table_names or [])

    def clear(self) -> None:
        self._cleared = True
        self.used_range.value = None

    def range(self, _cell: str) -> FakeRange:
        return FakeRange("dest")


class FakeSheets:
    def __init__(self, sheets: List[FakeSheet]) -> None:
        self._sheets = sheets

    def __iter__(self):
        return iter(self._sheets)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._sheets[key]
        for sheet in self._sheets:
            if sheet.name == key:
                return sheet
        raise KeyError(key)

    def add(self, name: str, after: FakeSheet):
        new_sheet = FakeSheet(name, used_range_value=None)
        self._sheets.append(new_sheet)
        return new_sheet


class FakeWorkbookApi:
    def __init__(self) -> None:
        self._pivot_caches = FakePivotCaches()

    def PivotCaches(self) -> FakePivotCaches:
        return self._pivot_caches


class FakeWorkbook:
    def __init__(self, sheets: List[FakeSheet]) -> None:
        self.sheets = FakeSheets(sheets)
        self.api = FakeWorkbookApi()


def _builder_with_table(
    workbook: FakeWorkbook,
    table_name: str = "Table1",
    destination_handling: DestinationHandling = DestinationHandling.FIND_OR_CREATE,
) -> PivotBuilder:
    return PivotBuilder(
        workbook=workbook,
        table_name=table_name,
        pivot_sheet_name="Pivot",
        pivot_table_name="PT_Test",
        pivot_top_left_cell="A3",
        row_fields=[RowField(name="Customer")],
        column_fields=[ColumnField(name="Category")],
        data_fields=[
            DataField(name="Qty", function=SummaryFunction.SUM, caption="Qty", number_format="0"),
            DataField(name="Total", function=SummaryFunction.SUM),
            DataField(name="Name", function=SummaryFunction.COUNT),
        ],
        destination_handling=destination_handling,
    )


def test_generate_pivot_happy_path() -> None:
    list_object = FakeListObject(["Customer", "Category", "Qty", "Total", "Name"])
    sheet = FakeSheet("Data", list_objects={"Table1": list_object}, used_range_value=None)
    pivot_sheet = FakeSheet("Pivot", used_range_value=None)
    workbook = FakeWorkbook([sheet, pivot_sheet])

    spec = _builder_with_table(workbook)
    spec.row_fields[0] = RowField(name="Customer", caption="Customer Name")
    spec.column_fields[0] = ColumnField(name="Category", caption="Item Category")
    spec.table_style = "PivotStyleMedium9"
    generate_pivot(spec)

    cache = workbook.api.PivotCaches().cache
    pivot_table = cache.table
    assert pivot_table is not None

    pf_customer = pivot_table.PivotFields("Customer")
    pf_category = pivot_table.PivotFields("Category")
    pf_qty = pivot_table.PivotFields("Qty")
    pf_total = pivot_table.PivotFields("Total")
    pf_name = pivot_table.PivotFields("Name")

    assert pf_customer.Orientation is not None
    assert pf_customer.Position == 1
    assert pf_customer.Caption == "Customer Name"
    assert pf_category.Position == 1
    assert pf_category.Caption == "Item Category"
    assert pf_qty.Function == XL_SUM
    assert pf_total.Function == XL_SUM
    assert pf_name.Function == XL_COUNT
    assert pivot_table.TableStyle2 == "PivotStyleMedium9"


def test_generate_pivot_force_clear() -> None:
    list_object = FakeListObject(["Customer", "Category", "Qty", "Total", "Name"])
    sheet = FakeSheet("Data", list_objects={"Table1": list_object}, used_range_value=None)
    pivot_sheet = FakeSheet("Pivot", used_range_value="data")
    workbook = FakeWorkbook([sheet, pivot_sheet])

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.EXISTING_FORCE_CLEAR)
    generate_pivot(spec)
    assert pivot_sheet._cleared is True


def test_validate_spec_inputs_errors() -> None:
    list_object = FakeListObject(["Customer", "Category", "Qty"])
    sheet = FakeSheet("Data", list_objects={"Table1": list_object}, used_range_value=None)
    workbook = FakeWorkbook([sheet])

    spec = _builder_with_table(workbook)
    spec.row_fields = spec.row_fields * 4
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)

    spec = _builder_with_table(workbook)
    spec.column_fields = spec.column_fields * 4
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)

    spec = _builder_with_table(workbook)
    spec.data_fields = []
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)

    spec = _builder_with_table(workbook)
    spec.data_fields = spec.data_fields * 7
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)

    spec = _builder_with_table(workbook)
    spec.destination_handling = None  # type: ignore[assignment]
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)

    spec = _builder_with_table(workbook)
    spec.destination_handling = "bad"  # type: ignore[assignment]
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)

    spec = _builder_with_table(workbook)
    spec.data_fields[0] = DataField(name="Qty", function="bad")  # type: ignore[arg-type]
    with pytest.raises(ValidationError):
        _validate_spec_inputs(spec)


def test_resolve_destination_sheet_paths() -> None:
    list_object = FakeListObject(["Customer", "Category", "Qty"])
    data_sheet = FakeSheet("Data", list_objects={"Table1": list_object}, used_range_value=None)
    existing_empty = FakeSheet("Pivot", used_range_value=None)
    existing_full = FakeSheet("PivotFull", used_range_value="x")
    workbook = FakeWorkbook([data_sheet, existing_empty, existing_full])

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.EXISTING_CLEAR)
    spec.pivot_sheet_name = "Pivot"
    assert _resolve_destination_sheet(spec).name == "Pivot"

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.EXISTING_NO_CLEAR)
    spec.pivot_sheet_name = "Pivot"
    assert _resolve_destination_sheet(spec).name == "Pivot"

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.FIND_OR_CREATE)
    spec.pivot_sheet_name = "PivotFull"
    with pytest.raises(DestinationError):
        _resolve_destination_sheet(spec)

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.NEW)
    spec.pivot_sheet_name = "Pivot"
    with pytest.raises(DestinationError):
        _resolve_destination_sheet(spec)

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.EXISTING_NO_CLEAR)
    spec.pivot_sheet_name = "Missing"
    with pytest.raises(DestinationError):
        _resolve_destination_sheet(spec)

    spec = _builder_with_table(workbook, destination_handling=DestinationHandling.FIND_OR_CREATE)
    spec.pivot_sheet_name = "NewPivot"
    created = _resolve_destination_sheet(spec)
    assert created.name == "NewPivot"


def test_find_list_object_and_column_names() -> None:
    list_object = FakeListObject(["A", "B"])
    sheet = FakeSheet("Data", list_objects={"Table1": list_object}, used_range_value=None)
    workbook = FakeWorkbook([sheet])

    assert _find_list_object(workbook, "Table1") is list_object
    assert _find_list_object(workbook, "Missing") is None
    assert _list_object_column_names(list_object) == ["A", "B"]
    assert _list_object_column_names(FakeListObjectBroken()) == []


def test_validate_and_get_table_errors() -> None:
    sheet = FakeSheet("Data", list_objects={}, used_range_value=None)
    workbook = FakeWorkbook([sheet])
    spec = _builder_with_table(workbook, table_name="Missing")

    with pytest.raises(ValidationError):
        _validate_and_get_table(spec)


def test_validate_unique_names_and_field_existence() -> None:
    _validate_unique_column_names(["A", "B", "C"])
    with pytest.raises(ValidationError):
        _validate_unique_column_names(["A", "a"])

    with pytest.raises(ValidationError):
        _validate_field_names_exist(
            [RowField(name="X")],
            [ColumnField(name="Y")],
            [DataField(name="Z", function=SummaryFunction.SUM)],
            ["A", "B"],
        )


def test_validate_pivot_table_name_unique() -> None:
    sheet = FakeSheet("Data", pivot_table_names=["PT_Test"], used_range_value=None)
    workbook = FakeWorkbook([sheet])

    with pytest.raises(ValidationError):
        _validate_pivot_table_name_unique(workbook, "PT_Test")

    class BrokenApi:
        def PivotTables(self):
            raise RuntimeError("broken")

    broken_sheet = FakeSheet("Broken", used_range_value=None)
    broken_sheet.api = BrokenApi()
    workbook = FakeWorkbook([broken_sheet])
    _validate_pivot_table_name_unique(workbook, "PT_Ok")


def test_summary_function_to_excel() -> None:
    assert _summary_function_to_excel(SummaryFunction.SUM) == XL_SUM
    assert _summary_function_to_excel(SummaryFunction.COUNT) == XL_COUNT
    assert _summary_function_to_excel(SummaryFunction.AVG) == XL_AVG
    with pytest.raises(ValidationError):
        _summary_function_to_excel("bad")  # type: ignore[arg-type]


def test_is_empty_value_paths() -> None:
    assert _is_empty_value(None) is True
    assert _is_empty_value("  ") is True
    assert _is_empty_value("x") is False
    assert _is_empty_value([None, ["", " "], []]) is True
