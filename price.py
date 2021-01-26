# Программа добавляющая цену с надбавкой в колонку "M"
from math import ceil
# Коэфф для товаров дешевле 800 рублей
CHEAPER_800 = 1.6

# Коэфф для товаров дороже 800 рублей
MORE_800 = 1.4


# Функция делает надбавку у цены и вписывает ее в колонку "M"
def add_extra_charge(ws):
    print('Производится надбавка цены')
    # Счетчик рядов
    row_counter = 1
    for cell in ws['L']:
        if cell.value == 'Цена':
            row_counter += 1
            continue
        elif cell.value < 800:
            price_with_charge = cell.value * CHEAPER_800
            result = ceil(price_with_charge)
            ws[f'M{row_counter}'].value = result
            row_counter += 1
        elif cell.value >= 800:
            price_with_charge = cell.value * MORE_800
            result = ceil(price_with_charge)
            ws[f'M{row_counter}'].value = result
            row_counter += 1
    print("Цена с надбавкой проставленна!")
