import xlwings as xw
import os


main_folder_path = "C:/Users/User/Desktop/Новая папка"


def process_files(folder_path):
    entries = os.listdir(folder_path)

    for entry in entries:
        entry_path = os.path.join(folder_path, entry)

        if os.path.isfile(entry_path):
            if entry.endswith(".xlsx") or entry.endswith(".xls"):
                wb = xw.Book(entry_path)
                sh2 = wb.sheets
                sh2.api.PrintOut(From=1, To=3, Copies=1)
                wb.close()

        elif os.path.isdir(entry_path):
            process_files(entry_path)


process_files(main_folder_path)
