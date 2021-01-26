from excel_file_search import *

EXCEL_FILE_NAME = excelfilename()

ROWS_TO_DELETE = [3, 8, 10, 10, 10, 10, 10, 10, 10, 10, 13, 13, 14, 14, 14, 15, 15, 15]


def delete_cols(wb, ws):
    row_counter = 0
    print('Удаление ненужных колонок')
    for i in range(1, len(ROWS_TO_DELETE) + 1):
        ws.delete_cols(ROWS_TO_DELETE[row_counter])
        row_counter += 1
        i += 1
    print('Удаление ненужных колонок завершенно!')
