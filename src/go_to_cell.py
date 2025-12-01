"""A script to figure out how to go to home (or any specific cell)"""
import xlwings as xw

wb = xw.Book()

ws = wb.sheets[0]

ws.range("E5").select()