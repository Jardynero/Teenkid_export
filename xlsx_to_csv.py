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
    print("Конвертирование выгрузки в csv формат")
    xl = pd.read_excel(excelfilename())
    xl.to_csv(CSV_FILE_NAME, index_label=None, header=True)
