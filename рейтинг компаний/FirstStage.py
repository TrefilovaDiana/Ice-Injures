# Первая стадия программы по составлению рейтинга компаний
# Определение рейтинга компаний в одном году

import math
import pylightxl as xl
import os, sys, inspect
# import lib.auxiliary
current_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir) 
from lib.auxiliary import *

# ------------------------- основной скрипт --------------------------

# выбираем файлик с падениями
file_path, file_name = openFile()
# открываем excel таблицу с падениями
fall_db = xl.readxl(fn=file_path)
sheet = fall_db.ws(ws=fall_db.ws_names[0])
# считываем адреса падений из excel-файла
fall_addresses = ReadAllFallAddresses(sheet) # [[индекс, город, улица, номер дома]]
# считываем данные обо всех домах и управ. компаниях
company_headers, companies = ReadAllCompanies("Организации.csv") # {id -> [данные компании]}
towns = ReadAllHouses('Дома.csv') # {город -> {улица -> {номер дома -> [company_id, кол-во жителей, категория]}}}

# ------------------------ собираем дома в компании -------------------------
# Проходим по каждому дому в городе, и формируем список вида: {id компании -> [[дома 1 кат], [дома 2 кат], [дома 3 кат]]}
houses_with_companies = dict()
for town in towns:
    for street in towns[town]:
        for house_num in towns[town][street]:
            # запоминаем данные о доме [company_id, кол-во жителей, категория]
            house_data = towns[town][street][house_num]
            # достаем id компании из данных о доме
            company_id = house_data[0]
            # вычленяем категорию дома
            category = house_data[2]
            # если в результирующем словарике нет компании, то вставляем пустые массивы в словарик
            if company_id not in houses_with_companies:
                houses_with_companies[company_id] = [[], [], [], []]
            # если категория у дома определена, то вставляем данные дома в словарик в соответствующую категорию
            if not category is None:
                houses_with_companies[company_id][category].append(house_data)


# ------------------- собираем травмы в компании -------------------------
# проходим по каждому месту падения, определяем адрес и дом, по дому определяем компанию
injuries_with_companies = dict() #{id компании -> [[падения 1 кат], [падения 2 кат], [падения 3 кат]]}
# injury - травма

# проходим по всем адресам падений
for address in fall_addresses:
    # достаем город
    town = address[1]
    # достаем улицу
    street = address[2]
    # достаем номер дома
    house = address[3]
    # ищем дом возле которого было это падение
    if town in towns and street in towns[town] and house in towns[town][street]:
        # если дом найден, то
        # сохраняем данные дома
        house_data = towns[town][street][house]
        # из данных дома достаем категорию и id компании
        category = house_data[2]
        company_id = house_data[0]
        # если в результирующем словарике еще нет такой компании, то вставляем пустые массивы
        if company_id not in injuries_with_companies:
            injuries_with_companies[company_id] = [[], [], [], []]
        # если категория определена, то вставляем данные о падении в компанию в категорию
        if not category is None:
            injuries_with_companies[company_id][category].append(address)

# ------------------------- Сводим все в один массив --------------------
table = []
# проходимся по каждой компании у которой были найдены травмы
for company_id in injuries_with_companies:
    # получаем список домов этой компании
    grouped_houses = houses_with_companies[company_id]
    # получаем список травм этой компании
    grouped_injures = injuries_with_companies[company_id]
    # получаем данные компании
    company = companies[company_id]
    # общее кол-во домов компании (по всем категориям)
    houses_count = 0
    # общее кол-во травм (по всем категориям)
    injures_count = 0
    # проходим по всем категориям и суммируем кол-во травм и домов
    for category in range(1, 4):
        houses_count += len(grouped_houses[category])
        injures_count += len(grouped_injures[category])
    # если дома есть и травмы тоже есть, то 
    if houses_count > 0 and injures_count > 0:
        # формируем строку таблицы: [данные компании, дома по категориям, травмы по категориям]
        row = [company, grouped_houses, grouped_injures]
        table.append(row)

# -------------------------- Сохраняем в эксель -------------------------

db = xl.Database()
db.add_ws("Sheet1")
sheet = db.ws("Sheet1")

# ---- записываем заголовки столбцов в excel-файл --- 
cur_col = 1
for header in company_headers:
    sheet.update_index(row=1, col=cur_col, val=header)
    cur_col += 1
# проходимся по всем категориям, и для каждой категории вставляем столбцы
for i in range(1, 4):
    sheet.update_index(row=1, col=cur_col, val="Кол-во травм(B)")
    sheet.update_index(row=1, col=cur_col + 1, val="Кол-во жителей(C)")
    sheet.update_index(row=1, col=cur_col + 2, val="Показатель травматизации(B/C)")
    sheet.update_index(row=1, col=cur_col + 3, val="Кол-во домов(E)")
    sheet.update_index(row=1, col=cur_col + 4, val="B/E")
    cur_col += 5

# записываем данные со второй строки экселя
cur_row = 2
# проходимся по каждой строке нашего сводного массива
for row in table:
    # данные о компании
    company = row[0]
    # дома по категориям
    grouped_houses = row[1]
    # травмы по категориям
    grouped_injures = row[2]
    # в начале записываем все данные о компании
    cur_col = 1
    for value in company:
        sheet.update_index(row=cur_row, col=cur_col, val=value)
        cur_col += 1
    # затем проходимся по каждой категории, считаем показатели и записываем их
    for category in range(1, 4):
        # определяем кол-во травм в категории
        injures_count = len(grouped_injures[category])
        # определеяем кол-во домов в категории
        houses_count = len(grouped_houses[category])
        # если дома есть и травмы тоже есть, то
        if houses_count > 0 and injures_count > 0:
            # считаем кол-во жителей во всех домах категории
            resident_count = 0
            # проходимся по каждому дому и суммируем кол-во жителей
            for house in grouped_houses[category]:
                resident_count += house[1]
            # показатель травматизма считаем
            injury_rate = injures_count / resident_count
            # кол-во травм / кол-во домов
            t = injures_count / houses_count
        else:
            # говорим что оценки посчитать нельзя
            resident_count = 'none'
            injury_rate = 'none'
            t = 'none'
        # записываем посчитанные показатели в эксель
        sheet.update_index(row=cur_row, col=cur_col, val=injures_count)
        sheet.update_index(row=cur_row, col=cur_col + 1, val=resident_count)
        sheet.update_index(row=cur_row, col=cur_col + 2, val=injury_rate)
        sheet.update_index(row=cur_row, col=cur_col + 3, val=houses_count)
        sheet.update_index(row=cur_row, col=cur_col + 4, val=t)
        cur_col += 5
    cur_row += 1

# удаляем файл result.xlsx если он уже существует
result_file_name = 'result' + file_name + '.xlsx'
try:
    os.remove(result_file_name)
except:
    pass
# сохраняем файл
xl.writexl(db, fn=result_file_name)
print("Готово")
