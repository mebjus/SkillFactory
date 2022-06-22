import numpy as np
import pandas as pd
import os
import requests
import pickle
import math
from pandas.api.types import CategoricalDtype

token = '25945DB021CBCB00A59775B430B5B8BC'

url = 'https://apitest.cityexpress.ru/v1/25945DB021CBCB00A59775B430B5B8BC/Calculate'

df = pd.DataFrame()
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.2F}'.format

for file in fullpaths:
    df1 = pd.read_excel(file, header=2, sheet_name=None)
    df1 = pd.concat(df1, axis=0).reset_index(drop=True)
    df = pd.concat([df, df1], axis=0)

price_dict = {}
price_freq = {}
price_freq_money = {}
price_freq_public = {}

filename = 'price.bin'
with open(filename, 'rb') as f:
    price_dict = pickle.load(open(filename, 'rb'))

###########

dirname = 'data/размер тн.xlsx'
df_tn = pd.read_excel(dirname)

df_tn['Дата'] = df_tn['Дата'].dt.to_period('D')
df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('D')

df = df.merge(df_tn, left_on='Дата Cоздания', right_on='Дата', how='left')
df['Размер'] = df['Размер'] + 1

# tn = 1.22
counter = 0

print(len(price_dict))
#
# ##########   ограничения на веса и стоимость
# df = df[df['Расчетный вес'] <= 0.25]
# df = df[df['Общая стоимость со скидкой'] > 220]

df = df[df['Общая стоимость со скидкой'] > 0]

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


########    выбор по своей географии

def ower_city(row):
    if str(row).upper() not in city_dict:
        return np.NAN
    else:
        return row


# df['Отправитель.Адрес.Город'] = df['Отправитель.Адрес.Город'].apply(ower_city)
# df['Получатель.Адрес.Город'] = df['Получатель.Адрес.Город'].apply(ower_city)
# df = df.dropna(how='any', axis=0)


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


df['Режим'] = df['Режим доставки'].apply(mod)
df['Режим'] = df['Режим'].astype('category')


def old(row):
    lst = (row['Отправитель.Адрес.Город'], row['Получатель.Адрес.Город'], row['вес'], row['Режим доставки'])
    if lst in price_dict.keys():
        if price_freq.get(lst) == None: price_freq[lst] = 0
        if price_freq_money.get(lst) == None: price_freq_money[lst] = 0
        price_freq[lst] += 1
        print(lst, ':', 'есть')
        if price_dict[lst] != 'нет тарифа': price_freq_money[lst] += row['Общая стоимость со скидкой']
        return price_dict[lst]
    else:
        return -1


def old_2(row):
    lst = (row['Отправитель.Адрес.Город'], row['Получатель.Адрес.Город'], row['вес'], row['Режим доставки'])
    if price_freq_public.get(lst) == None: price_freq_public[lst] = 0
    if lst in price_dict.keys():
        if row['price'] != 'нет тарифа':
            price_freq_public[lst] += row['price']
        return price_dict[lst]
    else:
        return -1


def tarif(row):
    global counter
    if row['Отправитель.Адрес.Город'] == 'Горское п, Выборгский р-н': return 0  # исключение ошибки
    lst = (row['Отправитель.Адрес.Город'], row['Получатель.Адрес.Город'], row['вес'], row['Режим доставки'])
    row['Режим доставки'] = row['Режим доставки'].upper().strip()
    if lst in price_dict.keys(): return price_dict[lst]
    if row['price'] == -1:
        params = {'cityFrom':       row['Отправитель.Адрес.Город'], 'cityTo': row['Получатель.Адрес.Город'],
                  'physicalWeight': row['вес'], 'name': row['Режим доставки'], 'quantity': '1', 'width': '5',
                  'height':         '5',
                  'length':         '5'}
        response = requests.get(url, params=params)
        for i in response.json()['Result']:
            if i['Name'].upper() == row['Режим доставки']:
                counter += 1
                print(counter, ':', lst, ':', round(i['TotalPrice'], 1))
                price_dict[lst] = i['TotalPrice']
                return i['TotalPrice']
    counter += 1
    price_dict[lst] = 'нет тарифа'
    print(counter, ':', lst, ':', 'нет тарифа')
    return 'нет тарифа'


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
    
    if row['Вид доставки'] == 'Международная': return round_custom(row['Расчетный вес'], 0.5)
    
    if row['Вид доставки'] == 'Местная' and row['Отправитель.Адрес.Город'] == 'Москва':
        if row['Режим'] == 'ПРАЙМ' or row['Режим'] == 'ЭКСПРЕСС':
            if row['Расчетный вес'] <= 1:
                return round_custom(row['Расчетный вес'], 0.25)
            elif row['Расчетный вес'] > 1:
                return round_custom(row['Расчетный вес'], 1)
    
    if row['Вид доставки'] == 'Областная' and row['Отправитель.Адрес.Город'] == 'Москва':
        if row['Расчетный вес'] <= 1:
            return round_custom(row['Расчетный вес'], 0.5)
        elif row['Расчетный вес'] > 1:
            return round_custom(row['Расчетный вес'], 1)
    
    if row['Вид доставки'] == 'Местная' and row['Отправитель.Адрес.Город'] == 'Санкт-Петербург':
        if row['Режим'] == 'ПРАЙМ' or row['Режим'] == 'ЭКСПРЕСС':
            if row['Расчетный вес'] <= 1:
                return round_custom(row['Расчетный вес'], 0.25)
            elif row['Расчетный вес'] > 1:
                return round_custom(row['Расчетный вес'], 1)
    
    if row['Вид доставки'] == 'Областная' and row['Отправитель.Адрес.Город'] == 'Санкт-Петербург':
        if row['Расчетный вес'] <= 1: return round_custom(row['Расчетный вес'], 1)
    
    if row['Вид доставки'] == 'Местная' and row['Отправитель.Адрес.Город'] != 'Москва' and row[
        'Отправитель.Адрес.Город'] != 'Санкт-Петербург':
        if row['Режим'] == 'ПРАЙМ' or row['Режим'] == 'ЭКСПРЕСС' or row['Режим'] == 'ОПТИМА':
            if row['Расчетный вес'] <= 1:
                return round_custom(row['Расчетный вес'], 1)
            elif row['Расчетный вес'] > 1:
                return round_custom(row['Расчетный вес'], 1)
    
    if row['Вид доставки'] == 'Областная' and row['Отправитель.Адрес.Город'] != 'Москва' and row[
        'Отправитель.Адрес.Город'] != 'Санкт-Петербург':
        if row['Расчетный вес'] <= 1:
            return round_custom(row['Расчетный вес'], 1)
        elif row['Расчетный вес'] > 1:
            return round_custom(row['Расчетный вес'], 1)
    return row['Расчетный вес']


df['Общая стоимость со скидкой'] = df['Общая стоимость со скидкой'].fillna(0)

df['вес'] = df.loc[:, ['Вид доставки', 'Расчетный вес', 'Режим доставки', 'Режим', 'Отправитель.Адрес.Город',
                       'Получатель.Адрес.Город']].apply(weight, axis=1)
df['price'] = df.loc[:, ['Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'вес', 'Режим доставки',
                         'Общая стоимость со скидкой']].apply(
    old, axis=1)

df['price'] = df['price'].fillna(0)

df['price'] = df.loc[:, ['Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'вес', 'Режим доставки',
                         'Общая стоимость со скидкой', 'price']].apply(
    old_2, axis=1)

df['price'] = df.loc[:, ['Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'вес', 'Режим доставки', 'price']].apply(
    tarif, axis=1)

print(len(price_dict))

print('Кол-во строк:', df.shape[0])


######### частота направлений
sorted_dict = {}
sorted_dict_money = {}
sorted_dict_money_public = {}

sorted_keys = sorted(price_freq, key=price_freq.get, reverse=True)
sorted_keys_money = sorted(price_freq_money, key=price_freq_money.get, reverse=True)
sorted_keys_money_public = sorted(price_freq_public, key=price_freq_public.get, reverse=True)

for w in sorted_keys:
    sorted_dict[w] = price_freq[w]
for w in sorted_keys_money:
    sorted_dict_money[w] = price_freq_money[w]
for w in sorted_keys_money_public:
    sorted_dict_money_public[w] = price_freq_public[w]

df_dict = pd.DataFrame(sorted_dict.items(), columns=['Кортеж', 'Кол отправлений'])
df_dict_money = pd.DataFrame(sorted_dict_money.items(), columns=['Кортеж', 'Продали'])

df_dict_money_public = pd.DataFrame(sorted_dict_money_public.items(), columns=['Кортеж', 'Паблик'])

df_dict['Кол отправлений'].dropna(inplace=True)
df_dict_money['Продали'].dropna(inplace=True)
df_dict_money_public['Паблик'].dropna(inplace=True)

df_dict = df_dict.merge(df_dict_money, how='left')
df_dict = df_dict.merge(df_dict_money_public, how='left')

df_dict = df_dict[df_dict['Кол отправлений'] > 0]

# ###### очистить "нет тарифа"
#
# # dict_new = {}
# # for i, j in price_dict.items():
# # 	if j != 'нет тарифа':
# # 		# print(i, j)
# # 		dict_new[i] = j
# # price_dict = dict_new.copy()
# # print(len(price_dict))
# # #######

f = open('price.bin', 'wb')
pickle.dump(price_dict, f)
f.close()

df1 = df[df['price'] == -1]
df2 = df[df['price'] == 'нет тарифа']

df = df[df['price'] != -1]
df = df[df['price'] != 'нет тарифа']

df['price'] = df['price'] * df['tn']

df = df[df['price'] > 0]

df['discount'] = (df['Общая стоимость со скидкой'] / df['price']) - 1

df_group = df.groupby(['ФО', 'Клиент'])[['Общая стоимость со скидкой', 'price']].agg(
    {'Общая стоимость со скидкой': 'sum', 'price': 'sum'})
df_group = df_group[df_group['Общая стоимость со скидкой'] > 0]
df_group = df_group.reset_index()
df_group['discount'] = (df_group['Общая стоимость со скидкой'] / df_group['price']) - 1

df = df.loc[:,
     ['ФО', 'Клиент', 'Номер отправления', 'Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'Расчетный вес',
      'вес', 'Режим доставки', 'Вид доставки',
      'Общая стоимость со скидкой', 'price', 'tn', 'discount']]

df2 = df2.loc[:,
      ['ФО', 'Клиент', 'Номер отправления', 'Отправитель.Адрес.Город', 'Получатель.Адрес.Город', 'Расчетный вес',
       'вес', 'Режим доставки', 'Вид доставки',
       'Общая стоимость со скидкой', 'price', 'tn']]

writer = pd.ExcelWriter('цены.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='итоги', startrow=1, index=False, header=False)
df2.to_excel(writer, sheet_name='не определен', startrow=1, index=False, header=False)
df_group.to_excel(writer, sheet_name='группировка', startrow=1, index=False, header=False)
df_dict.to_excel(writer, sheet_name='популярность', startrow=1, index=False)

workbook = writer.book

worksheet = writer.sheets['итоги']
worksheet2 = writer.sheets['популярность']
worksheet3 = writer.sheets['не определен']
worksheet4 = writer.sheets['популярность']

format = workbook.add_format({'border': 1, 'bg_color': '#E8FBE1', 'num_format': '#,##0'})
worksheet4.set_column('A:B', 50, format)
worksheet4.set_column('B:D', 15, format)

worksheet4.add_table(0, 0, df_dict.shape[0], 3, {'first_column': False, 'style': None})

header_format = workbook.add_format({
    'bold':       True,
    'text_wrap':  True,
    'valign':     'vcenter',
    'fg_color':   '#D7E4BC',
    'align':      'center_across',
    'num_format': '#,##0',
    'border':     1})

for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)
# for col_num, value in enumerate(df_dict.columns.values):
# 	worksheet4.write(0, col_num, value, header_format)
writer.save()
