import os
import pyexcel
import sys
import excel_file_search
from delete_cols import delete_cols
from xlsx_to_csv import convert_excel_to_csv
from articles import sortexcel, make_articles
from timeit import default_timer as timer
from openpyxl import load_workbook
import time
import datetime

# текущее время
localtime = time.localtime()

# Запустим таймер для проверки затраты скрипта
start_time = timer()

# Проверка на количество файлов в корневом каталоге
excel_file_search.number_of_xlsx_files()
# Определение название файла. По необзодимости конвертация в .xlsx или выход
excel_file_search.detect_file()
# Сортировка файла от "а до я" с помощью pandas
sortexcel()

# Инициализирую ексель файл
wb = load_workbook(excel_file_search.excelfilename())
ws = wb.active

# Функция производит изменения в Артикулах
make_articles(wb=wb, ws=ws)
# Функция удаляет ненужные колонки
delete_cols(wb=wb, ws=ws)

# Получение названия файла ексель
current_excel_file_name = excel_file_search.excelfilename()
# Константа - название для сохранения выгрузки.
XLSX_FILE_NAME = f'Выгрузка Teenkid от {datetime.date.today()}|{localtime.tm_hour}:{localtime.tm_min}.xlsx'

# Сохранение файла с расширением .xlsx
print('Соранение файла выгрузки с расширением .xlsx')
wb.save(filename=XLSX_FILE_NAME)
wb.close()
print('Файл выгрузки успешно сохранен!')

# Удаление дефолтной выгрузки(предоставленной поставщиков Teenkid
os.remove(current_excel_file_name)

# Конвертирую в csv
print('Конвертирование файла выгрузки в .csv формат для выгрузки на сайт')
convert_excel_to_csv()
print("Конвертирование файла в .csv формат успешно завершенно!")
print('На данный скрипт потраченно: {} времени'.format(timer() - start_time))
