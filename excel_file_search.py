import os
import pyexcel
import eel
import sys


# Функция проверяет, сколько файлов находится ексель находится в корневой папке
def number_of_xlsx_files():
    print("Проверка на количество файлов с расширением .xlsx в корневом каталоге")
    list_with_xlsx_files = []
    for file in os.listdir():
        if file.endswith('xlsx') is True:
            list_with_xlsx_files.append(file)
    if len(list_with_xlsx_files) > 1:
        # eel.toManyExcelFiles() # функция js для веб интерфейса вызывающая алерт
        print('В корневом каталоге находится более одного файла .xlsx! Необходимо оставить только один файл .xlsx')
        sys.exit()
    print('Проверка на количество файлов с расширением .xlsx пройдена!')


# Функция проверяет наличие файла ексель в корневой папке
def detect_file():
    print('Производится поиск имени .xlsx файла')
    for file in os.listdir():
        if file.endswith('.xls') is True:
            print("Найден файл с расширением .xls!")
            print("Производится конвертирование файла из .xls в .xlsx")
            pyexcel.save_book_as(file_name=file, dest_file_name='Таблица на импорт.xlsx')
            os.remove(file)
            excel_file_name = 'Таблица на импорт.xlsx'
            break
        elif file.endswith('.xlsx') is True:
            excel_file_name = file
            print('Нужный файл с .xlsx расширением обнаружен')
            break
        else:
            excel_file_name = None
    if excel_file_name is None:
        # eel.dontFindFileName()    # Функция JS для веб интерфейса которая вызывает алрет
        print('Файл с .xlsx расширением НЕ обнаружен!')
        sys.exit()
    return excel_file_name


# Функция возвращает название файла ексель
def excelfilename():
    for file in os.listdir():
        if file.endswith('.xlsx') is True:
            break
    return file
