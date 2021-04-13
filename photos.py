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


# Функция добавляет название фотографий с сайта черубино
def add_cherubino_photo(ws, os):
    for folder_name in os.listdir('img'):
        index_counter = 1
        for data in ws['A']:
            data = data.value
            if folder_name.find(data) == -1:
                index_counter += 1
                continue
            else:
                list_with_photos = str(os.listdir(f'img/{folder_name}'))
                list_with_photos = list_with_photos.replace('[', '')
                list_with_photos = list_with_photos.replace(']', '')
                list_with_photos = list_with_photos.replace("'", "")
                list_with_photos = list_with_photos.replace(' ', '')
                ws[f'C{index_counter}'].value = str(list_with_photos)
                index_counter += 1