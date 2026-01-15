import xlwings as xw
import polars as pl

def xlwings_polars():
    wb = xw.Book("Workbooks/xlwings_polars.xlsx")
    filename = wb.fullname
    print(filename)
    wb.close()

    wb_df = pl.read_excel(
        source=filename,
        sheet_name="Sheet1",
        engine='openpyxl'
    )
    print(wb_df)

if __name__ == "__main__":
    xlwings_polars()