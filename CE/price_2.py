from typing import Tuple, Any

import numpy as np
import pandas as pd
import os
import requests
import pickle
import math
from pandas.api.types import CategoricalDtype
import time

token = 'B6943005F57A8F24962C9DCC6537FC7F'
url = 'https://apitest.cityexpress.ru/v1/B6943005F57A8F24962C9DCC6537FC7F/Calculate'

df = pd.DataFrame()
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.0F}'.format

for file in fullpaths:
    df1 = pd.read_excel(file, header=2, sheet_name=None)
    df1 = pd.concat(df1, axis=0).reset_index(drop=True)
    df = pd.concat([df, df1], axis=0)

price_dict = {}

filename = 'data/price.bin'
with open(filename, 'rb') as f:
    price_dict = pickle.load(open(filename, 'rb'))

###########

dirname = 'data/размер тн.xlsx'
df_tn = pd.read_excel(dirname)

df_tn['Дата'] = df_tn['Дата'].dt.to_period('D')
df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('D')

df = df.merge(df_tn, left_on='Дата Cоздания', right_on='Дата', how='left')
df['Размер'] = df['Размер'] + 1
df.drop('Дата', axis=1, inplace=True)


dict_fo = {'СЗФО': ['ВЕЛИКИЙ НОВГОРОД', 'МУРМАНСК', 'ПЕТРОЗАВОДСК', 'СЫКТЫВКАР', 'САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК',
                    'КАЛИНИНГРАД'],
           'УФО':  ['КУРГАН', 'НИЖНЕВАРТОВСК', 'НОВЫЙ УРЕНГОЙ', 'СТЕРЛИТАМАК', 'МАГНИТОГОРСК', 'ОРЕНБУРГ', 'СУРГУТ',
                    'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК'],
           'ПФО':  ['ИЖЕВСК', 'ПЕНЗА', 'УЛЬЯНОВСК', 'ЧЕБОКСАРЫ', 'КИРОВ', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ', 'САМАРА',
                    'САРАТОВ', 'ТОЛЬЯТТИ'],
           'ЮФО':  ['НОВОРОССИЙСК', 'СИМФЕРОПОЛЬ', 'ПЯТИГОРСК', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
                    'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ'],
           'СФО':  ['БАРНАУЛ', 'НОВОКУЗНЕЦК', 'ТОМСК', 'УЛАН-УДЭ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
                    'КЕМЕРОВО'],
           'ДВФО': ['ВЛАДИВОСТОК', 'ХАБАРОВСК']}

city_dict = ['САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК',
             'КАЛИНИНГРАД', 'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ',
             'САМАРА',
             'САРАТОВ', 'ТОЛЬЯТТИ', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
             'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
             'КЕМЕРОВО', 'ВЛАДИВОСТОК', 'ХАБАРОВСК', 'МОСКВА']

begin = df.shape[0]

########    выбор по своей географии

def ower_city(row):
    if str(row).upper() not in city_dict:
        return np.NAN
    else:
        return row


# df['Отправитель.Адрес.Город'] = df['Отправитель.Адрес.Город'].apply(ower_city)
# df['Получатель.Адрес.Город'] = df['Получатель.Адрес.Город'].apply(ower_city)
# df = df.dropna(how='any', axis=0)

df.dropna(subset=['Общая стоимость со скидкой'], how='any', axis=0, inplace=True)

# print(df.shape[0]*100/begin)

df['Заказ.Клиент.Не применять топливную надбавку'] = df['Заказ.Клиент.Не применять топливную надбавку'].fillna(0)


def fuel(row):
    if row['Заказ.Клиент.Не применять топливную надбавку'] == 1:
        return 1
    else:
        return row['Размер']


df['tn'] = df.loc[:, ['Заказ.Клиент.Не применять топливную надбавку', 'Размер']].apply(fuel, axis=1)

df = df[~df['Режим доставки'].isin(
    ['ЭКСПРЕСС возврат документов', 'ЛОЖНЫЙ ВЫЗОВ', 'СКЛАД', 'ЭКСПРЕСС Груз', 'ВТОРИЧНАЯ ДОСТАВКА',
     'СИБИРСКИЙ ЭКСПРЕСС  Для физ.лиц', 'ВОЛЖСКИЙ ЭКСПРЕСС  склад-дверь до 0,5 кг', 'ЭКСПРЕСС B', 'ПРАЙМ А',
     'ЭКСПРЕСС А', 'ПРАЙМ B', 'ЭКОНОМ  склад-склад', 'ЮЖНЫЙ ЭКСПРЕСС  дверь-дверь', 'ЮЖНЫЙ ЭКСПРЕСС  дверь-склад',
     'ЭКСПРЕСС ДАЛЬНИЙ ВОСТОК  Для физ.лиц'])]


def ret(cell):  # столбец и ячейку передаю, возрат - округ
    for i in dict_fo.keys():
        if str(cell).upper() in dict_fo[i]:
            return i
    else:
        return 'ЦФО'


df['ФО'] = df['Заказ.Клиент.Подразделение.Адрес.Город'].apply(ret)
df['ФО'] = df['ФО'].astype('category')

cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)


def mod(arg):
    if arg.find('ЭКСПРЕСС') != -1:
        return 'ЭКСПРЕСС'
    elif arg.find('ПРАЙМ') != -1:
        return 'ПРАЙМ'
    elif arg.find('ОПТИМА') != -1:
        return 'ОПТИМА'
    else:
        return 'ПРОЧИЕ'

df['Режим доставки'] = df['Режим доставки'].astype(str)
df['Режим'] = df['Режим доставки'].apply(mod)
df['Режим'] = df['Режим'].astype('category')


def round_custom(num, step):
    return math.ceil(num / step) * step


def weight(row):
    if row['Вид доставки'] == 'Междугородная':
        if row['Режим доставки'] == 'ЭКОНОМ  склад-склад': return round_custom(row['Расчетный вес'], 1)
        if row['Расчетный вес'] <= 0.5:
            return 0.5
        elif (row['Расчетный вес'] > 0.5) and (row['Расчетный вес'] <= 20):
            return round_custom(row['Расчетный вес'], 0.5)
        elif (row['Расчетный вес'] > 20):
            return round_custom(row['Расчетный вес'], 1)

    elif row['Вид доставки'] == 'Международная':
        return round_custom(row['Расчетный вес'], 0.5)

    elif row['Вид доставки'] == 'Местная' and row['Отправитель.Адрес.Город'] == 'Москва':
        if row['Расчетный вес'] <= 0.5:
            return round_custom(row['Расчетный вес'], 0.25)
        elif row['Расчетный вес'] <= 1:
            return round_custom(row['Расчетный вес'], 0.5)
        elif row['Расчетный вес'] > 1:
            return round_custom(row['Расчетный вес'], 1)

    elif row['Вид доставки'] == 'Областная' and row['Отправитель.Адрес.Город'] == 'Москва':
        if row['Расчетный вес'] <= 1:
            return round_custom(row['Расчетный вес'], 0.5)
        else:
            return round_custom(row['Расчетный вес'], 1)

    elif row['Вид доставки'] == 'Местная' and row['Отправитель.Адрес.Город'] == 'Санкт-Петербург':
        if row['Расчетный вес'] <= 0.5:
            return round_custom(row['Расчетный вес'], 0.25)
        elif row['Расчетный вес'] <= 1:
            return round_custom(row['Расчетный вес'], 0.5)
        elif row['Расчетный вес'] > 1:
            return round_custom(row['Расчетный вес'], 1)

    elif row['Вид доставки'] == 'Областная' and row['Отправитель.Адрес.Город'] == 'Санкт-Петербург':
        return round_custom(row['Расчетный вес'], 1)

    elif row['Вид доставки'] == 'Местная' and row['Отправитель.Адрес.Город'] != 'Москва' and row[
        'Отправитель.Адрес.Город'] != 'Санкт-Петербург':
        return round_custom(row['Расчетный вес'], 1)

    elif row['Вид доставки'] == 'Областная' and row['Отправитель.Адрес.Город'] != 'Москва' and row[
        'Отправитель.Адрес.Город'] != 'Санкт-Петербург':
        return round_custom(row['Расчетный вес'], 1)
    return row['Расчетный вес']


df['Общая стоимость со скидкой'] = df['Общая стоимость со скидкой'].fillna(0)

df['вес'] = df.loc[:, ['Вид доставки', 'Расчетный вес', 'Режим доставки', 'Режим', 'Отправитель.Адрес.Город',
                       'Получатель.Адрес.Город']].apply(weight, axis=1)


# def check(row):
#     row = tuple(row)
#     if row in price_dict.keys():
#         return price_dict.get(row)

# df['price'] = df.loc[:, ['Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'вес', 'Режим доставки']].apply(check, axis=1)

##### очистим "нет тарифа"
# price_dict1={}
# for i in (price_dict.keys()):
#     if price_dict[i] != 'нет тарифа':
#         price_dict1[i] = price_dict[i]
# price_dict = price_dict1.copy()
#####

def get_a_price(row):
    if row['Отправитель.Адрес.Город'] == 'Горское п, Выборгский р-н': return 0  # исключение ошибки

    row['Режим доставки'] = row['Режим доставки'].upper().strip()
    row = tuple(row)

    if row in price_dict.keys():
        # print(row, '=', price_dict.get(row))
        return price_dict.get(row)

    params = {'cityFrom':row[0],'cityTo':row[1],'physicalWeight':row[2],'name':row[3],'quantity':'1','width':'5',
              'height':'5', 'length':'5'}
    response = requests.get(url, params=params)

    for i in response.json()['Result']:
        if i['Name'].upper() == row[3]:
            price_dict[row] = i['TotalPrice']
            print(row, '+', i['TotalPrice'])
            return i['TotalPrice']

    price_dict[row] = 'нет тарифа'
    print(row, '+', 'нет тарифа')
    return 'нет тарифа'


start_time = time.time()

df['price'] = df.loc[:, ['Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'вес', 'Режим доставки']].apply(get_a_price, axis=1)

print('размер словаря:', len(price_dict))
s = time.time() - start_time
print("---", s/60, " минут ---")
print('Кол-во строк:', df.shape[0])

f = open('data/price.bin', 'wb')
pickle.dump(price_dict, f)
f.close()


df = df.loc[:,
     ['ФО', 'Клиент', 'Номер отправления', 'Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'Расчетный вес',
      'вес', 'Режим доставки', 'Вид доставки',
      'Общая стоимость со скидкой', 'price', 'tn']]

df1 = df[df['price'] != 'нет тарифа']
df2 = df[df['price'] == 'нет тарифа']

df1['price'].astype('float64')
df1 = df1[df1['price'] > 0]
df1['price'] = df1['price'] * df1['tn']


df_group = df1.groupby(['ФО', 'Клиент'])[['Общая стоимость со скидкой', 'price']].agg(
    {'Общая стоимость со скидкой': 'sum', 'price': 'sum'}).reset_index()
df_group = df_group[df_group['price'] > 0]
df_group['discount'] = (df_group['Общая стоимость со скидкой'] / df_group['price']) - 1

writer = pd.ExcelWriter('data/цены.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name='итоги', index=True, header=True)
df2.to_excel(writer, sheet_name='нет тарифа', index=True, header=True)
df_group.to_excel(writer, sheet_name='группировка', index=True, header=True)
workbook = writer.book
writer.save()
