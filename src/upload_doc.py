"""Using tkinter to upload a doc and then run some manipulation with xlwing"""
import tkinter as tk
import xlwings as xw
from tkinter import filedialog
from pathlib import Path

print(Path(__file__).parent.parent) 

# Global Vars
my_book_location = ""
file_name = "My_Output_File"

# Functions()

def UploadAction(event=None):
    my_book_location = filedialog.askopenfilename()
    print("Selected: ", my_book_location)
    

def ManipulateBook(event=None):
    if my_book_location != "":
        print("No File Given")
        return
    my_book = xw.Book(my_book_location)
    my_worksheet = my_book.sheets[0]
    my_worksheet.range('A1').value = 'Hello, World'
    my_worksheet.range('A2:E20').value = 100
    my_book.save(Path(__file__).parent.parent._str + "/Workbooks/Output_Workbooks/" + file_name + ".xlsx")
    my_book.close()


# Build TK Window 
root = tk.Tk()

upload_button = tk.Button(root, text='Open', command=UploadAction)
upload_button.pack()

manipulate_button = tk.Button(root, text='Manipulate', command=ManipulateBook)
manipulate_button.pack()

root.mainloop()

