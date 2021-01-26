import pandas as pd
from excel_file_search import excelfilename
import time
import datetime

# Получаю массив с текущем временем
localtime = time.localtime()

# Константа с конструктором имени для csv файла
CSV_FILE_NAME = f'Выгрузка Teenkid от {datetime.date.today()}|{localtime.tm_hour}:{localtime.tm_min}.csv'


# Функция конвертирует xlsx в csv
def convert_excel_to_csv():
    print('Конвертирование файла выгрузки в .csv формат для выгрузки на сайт')
    xl = pd.read_excel(excelfilename())
    xl.to_csv(CSV_FILE_NAME, index_label=None, header=True)
    print("Конвертирование файла в .csv формат успешно завершенно!")
