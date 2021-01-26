# Программа удаляет в заголовках коды артикулов (для столбцов "F" и "G")


# Функуция убирает артикулы в начале строк в столбцах F и G
def title_correction(ws):
    print('Коррекция названий товаров')
    # Счетчик рядов
    row_counter = 1
    for cell in ws['A']:
        if cell.value == 'Артикул':
            row_counter += 1
            continue
        current_value = ws[f'E{row_counter}'].value
        ws[f'E{row_counter}'].value = current_value[len(cell.value)::]
        current_value = ws[f'F{row_counter}'].value
        ws[f'F{row_counter}'].value = current_value[len(cell.value)::]
        row_counter += 1
    # Убираю пробелы в начале строк столбца F

    # Счетчик рядов
    row_counter = 1
    for cell in ws['E']:
        current_value = cell.value
        current_value = current_value.lstrip()
        ws[f'E{row_counter}'].value = current_value
        row_counter += 1
    # Убираю пробелы в начале строк столбца G

    # Счетчик рядов
    row_counter = 1
    for cell in ws['F']:
        current_value = cell.value
        current_value = current_value.lstrip()
        ws[f'F{row_counter}'].value = current_value
        row_counter += 1
    print("Коррекция названий завершенна!")
