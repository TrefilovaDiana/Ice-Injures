import csv
import codecs
import math
import tkinter as tk
from tkinter import filedialog


# прочитать все адреса падений с листа excel
def ReadAllFallAddresses(sheet):
    # выделяем колонку с городом
    town_column = sheet.col(col=2)
    # выделяем колонку с адресом
    address_column = sheet.col(col=3) # [address]
    # убираем из выделенных колонок самую первую ячейку, потому что там название колонки
    town_column.pop(0)
    address_column.pop(0)
    # собираем массив индексов для нумерации каждой строки
    indices = [i for i in range(2, len(address_column))]
    # разделяем адрес на отдельно улицу, отдельно дом
    address_column = list(map(lambda x: [y.strip() for y in x.split(',')], address_column)) # [[street, house]]
    # объединяем списки в один список: список индексов, список городов, список адресов => [[индекс, город, улица, дом]]
    # список индексов, список городов, список адресов => [[индекс, город, [улица, дом]]]
    fall_addresses = zip(indices, town_column, address_column)
    # [[индекс, город, [улица, дом]]] => [[индекс, город, улица, дом]]
    fall_addresses = list(map(lambda x: [x[0], x[1]] + x[2], list(fall_addresses)))
    return fall_addresses


# прочитать все данные о компаниях
def ReadAllCompanies(file_path):
    # открываем файлик - таблицу с компаниями
    with codecs.open(file_path, 'r', 'utf_8_sig') as file:
        # создаем читателя csv таблицы. Разделитель значений - ;
        companies_reader = csv.reader(file, delimiter=";")
        # считываем первую строку - заголовки в переменную
        company_header = next(companies_reader, None)
        # убираем первый столбик - id
        company_header.pop(0)
        # создаем пустой словарь компаний. Словарь нужен, чтобы быстро искать компанию по id
        companies = {}
        # проходимся по каждой компании в файле и читаем строку
        for company in companies_reader:
            # запоминаем id
            id = company[0]
            # выкидываем id из компании
            company.pop(0)
            # записываем в словарь данные о компании
            companies[id] = company
    return company_header, companies # возвращает два значения: заголовки таблицы и строки таблицы


# прочитать все данные о домах
def ReadAllHouses(file_path):
    # открываем файлик - таблицу с домами
    with codecs.open(file_path, "r", 'utf_8_sig') as file:
        # создаем читатетя csv файла. Разделитель - ;
        houses_reader = csv.reader(file, delimiter=";")
        # считываем первую строку файла - там названия столбиков
        next(houses_reader, None)
        # создаем словарик для домов. {город -> {улица -> {номер дома -> данные о доме}}}
        towns = {}
        
        # считываем данные обо всех домах
        for house_data in houses_reader:
            # запоминаем город
            town = house_data[10]
            # запоминаем улицу
            street = house_data[12]
            # запоминаем номер дома
            house = house_data[13]
            # запоминаем id управляющей компании
            company_id = house_data[19]
            # ---- определяем кол-во жителей в доме и категорию дома ----
            # если указана жилплощадь в доме
            if house_data[35]:
                # считываем жилплощадь
                area_residential = float(house_data[35].replace(',', '.'))  # жилплощадь в доме
                # определяем кол-во человек в доме. Math.ceil - округлить число в большую сторону
                residents_count = int(math.ceil(area_residential / 15))  # кол-во человек в доме в среднем
                # Определяем дом в категорию
                if residents_count >= 1000:
                    category = 1
                elif residents_count >= 200:
                    category = 2
                else:
                    category = 3
            else: # если жилплощадь не указана, то говорим, что кол-во жителей и категория не указана
                residents_count = None
                category = None

            # если управляющая компания найдена, то вставляем данные о доме в словарь
            if company_id:
                # формируем данные о доме
                house_data = [company_id, residents_count, category]
                # если в словаре городов город еще не был указан, то вставляем новый город в словарь
                if town not in towns:
                    towns[town] = {street: {house: house_data}}
                # если город указан, но в словаре улиц не указана улица, то вставляем новую улицу в словарь 
                elif street not in towns[town]:
                    towns[town][street] = {house: house_data}
                # если есть и город и улица, то вставляем данные о доме
                else:
                    towns[town][street][house] = house_data
    return towns


# выбрать файл на жестком диске
def openFile():
    # служебный код
    root = tk.Tk()
    root.withdraw()
    # открываем окно для выбора файлика
    filepath = filedialog.askopenfilename()
    # определяем название файла
    filename = filepath.split('/')[-1].split('.')[0]
    return filepath, filename # возвращает два значения: путь до файла и название файла
