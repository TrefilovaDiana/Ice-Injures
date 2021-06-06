import os
import sys, inspect
import pylightxl as xl
# import lib.auxiliary
current_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir) 
from lib.auxiliary import *

# ---------------------------------- Основной скрипт -----------------------------------

# открываем окно в котором выбираем файлик который будем заполнять
file_path, file_name = openFile()

# открываем excel таблицу с падениями
fall_db = xl.readxl(fn=file_path)
sheet = fall_db.ws(ws=fall_db.ws_names[0])
# считываем адреса падений из excel-файла
fall_addresses = ReadAllFallAddresses(sheet)    # [[индекс, город, улица, номер дома]]
# считываем данные обо всех домах и управ. компаниях
company_headers, companies = ReadAllCompanies("Организации.csv") # {id -> [данные компании]}
towns = ReadAllHouses('Дома.csv')  # {город -> {улица -> {номер дома -> [company_id, кол-во жителей, категория]}}}

# ------------------------ перебираем все адреса и находим для каждого управ. компанию -------------------------
for address in fall_addresses:
    # находим город
    town = address[1]
    # находим улицу
    street = address[2]
    # находим номер дома
    house = address[3]
    # находим дом возле которого было падение
    if town in towns and street in towns[town] and house in towns[town][street]:
        address.append(towns[town][street][house]) # [индекс, город, улица, номер дома] + [company_id, кол-во жителей, категория] = [индекс, город, улица, номер дома, [company_id, кол-во жителей, категория]]
    else:
        address.append(['no company', 'None', 'None'])

# -------------------- Записываем результат в новый excel-файл -----------------------------------

# записываем заголовки столбцов в excel-файл
cur_col = 9
for header in company_headers:
    sheet.update_index(row=1, col=cur_col, val=header)
    cur_col += 1
sheet.update_index(row=1, col=cur_col, val="Кол-во жителей")
sheet.update_index(row=1, col=cur_col + 1, val="Категория")

#идем по адресам, смотрим какая у каждого адреса компания и записываем их в файл
for address in fall_addresses:
    # индекс падения
    row = address[0]
    # определяем id компании
    company_id = address[4][0]
    # кол-во жителей
    residents_count = address[4][1]
    # категория
    category = address[4][2]
    # столбик с которого будем вставлять значения
    cur_col = 9
    # если компания найдена
    if company_id != 'no company':
        # ищем данные компании по id
        company_data = companies[company_id]
        # каждое значение из данных компании записываем в excel
        for value in company_data:
            sheet.update_index(row=row, col=cur_col, val=value)
            cur_col += 1
        # если кол-во жителей и категория определены, то вставляем и их
        if not residents_count is None and not category is None:
            sheet.update_index(row=row, col=cur_col, val=residents_count)
            sheet.update_index(row=row, col=cur_col + 1, val=category)


result_file_name = 'result' + file_name + '.xlsx'
try:
    os.remove(result_file_name)
except:
    pass
# сохраняем файл
xl.writexl(fall_db, fn=result_file_name)
print("Готово")
