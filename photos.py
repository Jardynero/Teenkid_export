# Файл нужен для добавления данных в колонку Photos

# Расширение файла у фотографий
IMAGE_EXTENTION = '.jpg'


# Функция проставляет названия фотографий в колонку "C"
def add_photo_titles(ws):
    # Счетчик рядов
    row_counter = 1
    print('Проставление фотогирафий товаров')
    for cell in ws['A']:
        if cell.value == 'Артикул':
            row_counter += 1
            continue
        else:
            article = cell.value
            delete_spaces = article.replace(' ', '')
            article = delete_spaces
            ws[f'C{row_counter}'].value = article + IMAGE_EXTENTION
            row_counter += 1
    print('Проставление фотографий товаров завершенно!')
