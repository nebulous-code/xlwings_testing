"""
Example wrapper that builds a PivotSpec and generates a pivot table.

This script owns the Excel app and workbook lifecycle for demonstration purposes.
"""

from __future__ import annotations

import xlwings as xw

from pivot_util import (
    ColumnFieldSpec,
    DataFieldSpec,
    DestinationHandling,
    PivotSpec,
    RowFieldSpec,
    SummaryFunction,
)


def main() -> None:
    workbook_path = r"C:\Users\nlicalsi\Documents\Code\xlwings_testing\Workbooks\pivot_table_example.xlsx"

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(workbook_path)

        # Build the pivot specification.
        spec = PivotSpec(
            workbook=wb,
            table_name="Table1",
            pivot_sheet_name="Pivot",
            pivot_table_name="PT_By_Customer",
            pivot_top_left_cell="A3",
            row_fields=[
                RowFieldSpec(name="Customer"),
                RowFieldSpec(name="Category"),
            ],
            column_fields=[],
            data_fields=[
                DataFieldSpec(
                    name="Name",
                    function=SummaryFunction.COUNT,
                    caption="Orders",
                    number_format="0",
                ),
                DataFieldSpec(
                    name="Qty",
                    function=SummaryFunction.SUM,
                    caption="Total Qty",
                    number_format="0",
                ),
                DataFieldSpec(
                    name="Total",
                    function=SummaryFunction.SUM,
                    caption="Total Spent",
                    number_format="$#,##0.00",
                ),
            ],
            destination_handling=DestinationHandling.NEW,
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
