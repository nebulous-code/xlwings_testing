"""
Create a PivotTable from an existing Excel Table (ListObject) using xlwings (Windows/COM).

Assumptions:
- The source data is an Excel Table (Insert -> Table) with a known name, e.g. "All_Rows"
- You want a pivot on a new (or existing) sheet
- Windows Excel (COM) is available
"""

from __future__ import annotations

import xlwings as xw


# Excel PivotField orientation constants (COM)
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_DATA_FIELD = 4

# Excel summary function constants (COM)
XL_SUM = -4157
XL_COUNT = -4112


def create_pivot_from_table(
    workbook_path: str,
    table_name: str,
    pivot_sheet_name: str = "Pivot",
    pivot_table_name: str = "PivotTable1",
    pivot_top_left_cell: str = "A3",
) -> None:
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(workbook_path)

        # --- Find the ListObject (Excel Table) by name across all sheets ---
        list_object = None
        source_sheet = None
        for sht in wb.sheets:
            try:
                lo = sht.api.ListObjects(table_name)  # type: ignore[attr-defined]
                # If it exists, ListObjects(table_name) returns a COM object (no exception)
                list_object = lo
                source_sheet = sht
                break
            except Exception:
                continue

        if list_object is None or source_sheet is None:
            raise ValueError(f"Table '{table_name}' not found in workbook.")

        # The pivot cache source range should be the Table's range (includes headers)
        source_range = list_object.Range  # COM Range

        # --- Ensure pivot sheet exists (create or clear) ---
        try:
            pivot_sheet = wb.sheets[pivot_sheet_name]
        except Exception:
            pivot_sheet = wb.sheets.add(pivot_sheet_name, after=wb.sheets[-1])
            

        # Optional: clear existing content
        pivot_sheet.clear()

        # --- Build pivot cache + pivot table ---
        # SourceType=1 => xlDatabase
        pivot_cache = wb.api.PivotCaches().Create(1, source_range)  # type: ignore[attr-defined]

        dest = pivot_sheet.range(pivot_top_left_cell).api  # COM Range
        pivot_table = pivot_cache.CreatePivotTable(dest, pivot_table_name)

        # --- Configure fields for your table ---
        # Rows: Category
        # Values: Count of Name (orders) + Sum of Total (category total)

        # Row field
        pf_category = pivot_table.PivotFields("Category")
        pf_category.Orientation = XL_ROW_FIELD
        pf_category.Position = 1

        # Values: Count of orders (count Name)
        pf_count = pivot_table.PivotFields("Name")
        pf_count.Orientation = XL_DATA_FIELD
        pf_count.Function = XL_COUNT

        # Values: Sum of Total
        pf_sum = pivot_table.PivotFields("Total")
        pf_sum.Orientation = XL_DATA_FIELD
        pf_sum.Function = XL_SUM

        # Optional: nicer formatting
        pivot_table.ShowTableStyleRowStripes = True
        pivot_table.TableStyle2 = "PivotStyleMedium9"

        wb.save()

    finally:
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


if __name__ == "__main__":
    create_pivot_from_table(
        workbook_path=r"C:\Users\nlicalsi\Documents\Code\xlwings_testing\Workbooks\pivot_table_example.xlsx",
        table_name="Table1",
        pivot_sheet_name="Pivot",
        pivot_table_name="PT_All_Rows",
        pivot_top_left_cell="A3",
    )
