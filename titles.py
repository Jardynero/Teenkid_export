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

    # Добавляю артикулы в конец строк колонки "E"
    # Счетчик рядов
    row_counter = 1
    for cell in ws['A']:
        ws[f'E{row_counter}'].value = str(ws[f'E{row_counter}'].value + " " + cell.value)
        row_counter += 1


# Функция убирает оставшиеся скобочки вначале строк столбца "наименование"
def second_title_correction_e(ws):
    column_e = []
    for cell in ws['E']:
        column_e.append(cell.value)

    index_counter = 0
    for title in column_e:
        if title.find('(') >= 0:
            final_title = title[(title.find(' ') + 1)::]
            column_e[index_counter] = final_title
            index_counter += 1
        else:
            index_counter += 1

    index_counter = 0
    for title in column_e:
        if title.find('-') >= 0:
            final_title = title[(title.find(' ') + 1)::]
            column_e[index_counter] = final_title
            index_counter += 1
        else:
            index_counter += 1

    row_counter = 1
    for title in column_e:
        ws[f'E{row_counter}'].value = title
        row_counter += 1


# Функция убирает оставшиеся скобочки вначале строк столбца "полное наименование"
def second_title_correction_f(ws):
    column_e = []
    for cell in ws['F']:
        column_e.append(cell.value)

    index_counter = 0
    for title in column_e:
        if title.find('(') == 0:
            final_title = title[(title.find(' ') + 1)::]
            column_e[index_counter] = final_title
            index_counter += 1
        else:
            index_counter += 1

    index_counter = 0
    for title in column_e:
        if title.find('-') == 0:
            final_title = title[(title.find(' ') + 1)::]
            column_e[index_counter] = final_title
            index_counter += 1
        else:
            index_counter += 1

    row_counter = 1
    for title in column_e:
        ws[f'F{row_counter}'].value = title
        row_counter += 1
    print("Коррекция названий завершенна!")
