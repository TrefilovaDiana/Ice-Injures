import os
import pylightxl as xl
# названия файликов откуда надо брать данные - результат первой стадии
paths = ['result2015.xlsx', 'result2016.xlsx', 'result2017.xlsx', 'result2018.xlsx', 'result2019.xlsx']

# считать все компании из одного файлика за один год
def ReadCompaniesTop(path, year):
    # открываем эксель
    fall_db = xl.readxl(fn=path)
    # открываем лист
    sheet = fall_db.ws(ws=fall_db.ws_names[0])
    row_index = 1
    companies = dict() #{инн -> данные о компаниях}
    # массив для заголовков
    headers = []
    # проходимся по всем строкам в экселе
    for row in sheet.rows:
        # если строка первая, то считываем заголовки
        if row_index == 1:
            headers = row
            # в начало добавляем столбец Год
            headers.insert(0, 'year')
        else:
        # если строка не первая, то считываем данные компании
            inn = row[4] # определяем инн
            companies[inn] = row # вставляем данные о компании в словарь
            companies[inn].insert(0, year) # добавляем данные о годе
        row_index += 1
    return headers, companies # возвращает два значения: заголовки и данные о компаниях

# считываем первый файлик
headers, intersect_companies = ReadCompaniesTop(paths[0], 2015)
# меняем формат словаря. Было: {инн -> данные компании}. Стало: {инн -> [данные за 2015, данные за 2016, данные за 2017, данные за 2018, данные за 2019]}
for inn in intersect_companies:
    intersect_companies[inn] = [intersect_companies[inn]]

# проходим по всем оставшимся файликам
for i in range(1, len(paths)):
    # считываем компании из файла
    _, companies = ReadCompaniesTop(paths[i], 2015 + i)
    # формируем список инн, которые должны быть удалены из пересечения
    inn_to_delete = []
    # проходим по всем инн в пересечении
    for inn in intersect_companies:
        # если инн не найдено в очередном считанном файле, то инн удаляется из пересечения
        if inn not in companies:
            inn_to_delete.append(inn)
        else: # иначе заносим данные за год в результирующий словарь
            intersect_companies[inn].append(companies[inn])
    # удаляем из пересечения все инн которые не были найдены в очередном файлике
    for inn in inn_to_delete:
        del intersect_companies[inn]

db = xl.Database()
db.add_ws("sheet1")
sheet = db.ws('sheet1')

#--------------- заголовки -----------------
cur_col = 1
for header in headers:
    sheet.update_index(row=1, col= cur_col, val = header)
    cur_col += 1

#------------ Значения ----------------
cur_row = 2
for inn in intersect_companies:
    # получаем данные о компании за все 5 лет
    company = intersect_companies[inn]
    # для каждого года в данных
    for year in company:
        cur_col = 1
        # записываем все значения за один год
        for value in year:
            sheet.update_index(row=cur_row, col=cur_col, val=value)
            cur_col += 1
        cur_row += 1


# удаляем файл result.xlsx если он уже существует
try:
    os.remove('result.xlsx')
except:
    pass
# сохраняем файл
xl.writexl(db=db, fn='result.xlsx')
print("Готово")



