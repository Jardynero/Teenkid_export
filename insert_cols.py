# Файл нужен для добавления новых необходимых колонок в файл xlsx
COLS_NUMBERS = [3, 5, 9, 13, 20, 21]


# Функция добавляет колонки в таблицу
def inserting_cols(ws):
    print('Добавление новых столбцов в таблицу')
    # счетчик индекса
    index_counter = 0
    for col_num in range(1, len(COLS_NUMBERS) + 1):
        ws.insert_cols(COLS_NUMBERS[index_counter])
        index_counter += 1
    print(f'{len(COLS_NUMBERS)} новых столбцов, успешно добавленно!')


# Функция даст имена новым столбцам
def name_columns(ws):
    print('Именование новых столбцов')
    ws['C1'].value = 'Photos'
    ws['E1'].value = 'Categories'
    ws['I1'].value = 'Filter color'
    ws['M1'].value = 'Price with extra charge'
    ws['T1'].value = 'Up Sale'
    ws['U1'].value = 'Metka'
    print("Именование завершенно!")
