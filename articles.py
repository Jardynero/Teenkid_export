from excel_file_search import excelfilename
import pandas as pd
import os

# Константа - путь к ексель файлу
EXCEL_FILE_NAME = excelfilename()


# Функция сортирует ексель что бы все артикулы шли по порядку
def sortexcel():
    print('Сортировка таблицы в необходимом порядке')
    xl = pd.ExcelFile(EXCEL_FILE_NAME)
    df = xl.parse('TDSheet')
    df = df.sort_values(by='Артикул')
    # Сохраняем эксель
    writer = pd.ExcelWriter(EXCEL_FILE_NAME)
    df.to_excel(writer, index=False)
    writer.save()
    writer.close()


def make_articles(wb, ws):
    # Счетчик строк
    counter = 1

    # Убираю все данные в скобочках
    for cell in ws['A']:
        data_in_cell = cell.value
        if data_in_cell.find('(') < 0:
            counter += 1
            continue
        else:
            ws['A{}'.format(str(counter))].value = data_in_cell[:(data_in_cell.find('(') - 1)]
            counter += 1

    # Счетчик строк
    counter = 1

    # Убираю все префиксы "-"
    for cell in ws['A']:
        data_in_cell = cell.value
        if data_in_cell.find('-') < 0:
            counter += 1
            continue
        else:
            ws['A{}'.format(str(counter))].value = data_in_cell[:data_in_cell.find('-')]
            counter += 1

    # Счетчик строк
    counter = 1

    # Пустой список для добавления артикулов из екселя
    empty_list = []

    # Добавляю артикулы из екселя
    for cell in ws['A']:
        empty_list.append(cell.value)

    # Счетчик строк
    counter = 0
    counter_row = 1
    # Убираю пробелы вконце каждого артикуля
    for article in empty_list:
        empty_list[counter] = article.rstrip()
        ws['A{}'.format(str(counter_row))].value = article
        counter += 1
        counter_row += 1

    # Список который будет хранить списки с артикулами
    list_with_all_articles = []

    # Список который будет хранить одинаковые артикулы
    list_with_articles = []

    # Список со всеми артикулами из Екселя
    all_articles_from_excel = []

    # Добавляю все артикулы в список
    for cell in ws['A']:
        all_articles_from_excel.append(cell.value)

    # Формирую списки из  дублирующихся артикулов
    for current_article in list(dict.fromkeys(all_articles_from_excel)):
        list_with_articles = []
        counter = all_articles_from_excel.count(current_article)
        for i in range(1, (counter + 1)):
            list_with_articles.append(current_article)
        list_with_all_articles.append(list_with_articles)

    # Добавляю префиксы к артикулам
    for sub_list in list_with_all_articles:
        counter = 0
        article_counter = 1
        for article in sub_list:
            sub_list[counter] = str(article) + str('-{}'.format(str(article_counter)))
            counter += 1
            article_counter += 1

    # Счетчик строк
    counter = 1

    # Добавляю артикулы с префиксами в столбик "В" в екселе
    for sub_list in list_with_all_articles:
        for article in sub_list:
            ws['B{}'.format(str(counter))].value = article
            counter += 1
