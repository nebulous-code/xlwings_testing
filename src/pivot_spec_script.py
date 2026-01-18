"""
Example wrapper that builds a PivotBuilder and generates a pivot table.

This script owns the Excel app and workbook lifecycle for demonstration purposes.
"""

from __future__ import annotations

import os

import xlwings as xw

from pivot_util import (
    DataField,
    DestinationHandling,
    PivotBuilder,
    RowField,
    SummaryFunction,
)


def _ensure_workbook_closed(workbook_path: str) -> None:
    """
    Fail fast if the workbook is already open in Excel.

    Args:
        workbook_path: Full path to the workbook on disk.

    Returns:
        None

    Raises:
        RuntimeError: If the workbook is already open.
    """

    target = os.path.abspath(workbook_path).lower()
    for app in xw.apps:
        for book in app.books:
            try:
                fullname = os.path.abspath(book.fullname).lower()
            except Exception:
                continue
            if fullname == target:
                raise RuntimeError(
                    f"Workbook is already open in Excel: {workbook_path}"
                )


def main() -> None:
    workbook_path = r"C:\Users\nlicalsi\Documents\Code\xlwings_testing\Workbooks\pivot_table_example.xlsx"

    _ensure_workbook_closed(workbook_path)

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(workbook_path)

        # Build the pivot specification.
        spec = PivotBuilder(
            workbook=wb,
            table_name="Table1",
            pivot_sheet_name="Pivot",
            pivot_table_name="PT_By_Customer",
            pivot_top_left_cell="A3",
            row_fields=[
                RowField(name="Customer"),
                RowField(name="Category"),
            ],
            column_fields=[],
            data_fields=[
                DataField(
                    name="Name",
                    function=SummaryFunction.COUNT,
                    caption="Orders",
                    number_format="0",
                ),
                DataField(
                    name="Qty",
                    function=SummaryFunction.SUM,
                    caption="Total Qty",
                    number_format="0",
                ),
                DataField(
                    name="Total",
                    function=SummaryFunction.SUM,
                    caption="Total Spent",
                    number_format="$#,##0.00",
                ),
            ],
            destination_handling=DestinationHandling.FIND_OR_CREATE,
            table_style="PivotStyleMedium9",
            show_row_stripes=True,
        )

        # Validate and generate the pivot table.
        spec.generate_pivot()

        wb.save()
    finally:
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


if __name__ == "__main__":
    main()
