import xlwings as xw
import os 

def copy_tables():
    book1 = xw.Book(os.getcwd() + "/Workbooks/book_with_tables.xlsx")
    book1.sheets[0].tables["Table1"].data_body_range.copy()
    book1.sheets[0].tables["Table2"].data_body_range.paste("values")

if __name__ == "__main__":
    copy_tables()