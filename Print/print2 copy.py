import xlwings as xw
import xlwings.constants

wb=xw.Book("C:/Users/User/Desktop/Новая папка/ТТН АРСК.xls")
sh2=wb.sheets
sh2.api.PrintOut(From=1, To=3, Copies=1)