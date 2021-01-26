import os
import pyexcel
import eel
import sys


# Функция проверяет, сколько файлов находится ексель находится в корневой папке
def number_of_xlsx_files():
    list_with_xlsx_files = []
    for file in os.listdir():
        if file.endswith('xlsx') is True:
            list_with_xlsx_files.append(file)
    if len(list_with_xlsx_files) > 1:
        eel.toManyExcelFiles()
        sys.exit()


# Функция проверяет наличие файла ексель в корневой папке
def detect_file():
    for file in os.listdir():
        if file.endswith('.xls') is True:
            pyexcel.save_book_as(file_name=file, dest_file_name='Таблица на импорт.xlsx')
            os.remove(file)
            excel_file_name = 'Таблица на импорт.xlsx'
            break
        elif file.endswith('.xlsx') is True:
            excel_file_name = file
            break
        else:
            excel_file_name = None
    if excel_file_name is None:
        eel.dontFindFileName()
        sys.exit()
    return excel_file_name


# Функция возвращает название файла ексель
def excelfilename():
    for file in os.listdir():
        if file.endswith('.xlsx') is True:
            break
    return file
