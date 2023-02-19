###### процент своей или агентской географии


import numpy as np
import pandas as pd
import os
from pandas.api.types import CategoricalDtype

df = pd.DataFrame()
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.0F}'.format

for file in fullpaths:
	if file == 'data/kis/.DS_Store': os.remove('data/kis/.DS_Store')
	df1 = pd.read_excel(file, header=2, sheet_name=None)
	df1 = pd.concat(df1, axis=0).reset_index(drop=True)
	df = pd.concat([df, df1], axis=0)

df['Заказ.Клиент.Не применять топливную надбавку'] = df['Заказ.Клиент.Не применять топливную надбавку'].fillna(0)

dict_fo = {'СЗФО': ['ВЕЛИКИЙ НОВГОРОД', 'МУРМАНСК', 'ПЕТРОЗАВОДСК', 'СЫКТЫВКАР', 'САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК',
                    'КАЛИНИНГРАД'],
           'УФО': ['КУРГАН', 'НИЖНЕВАРТОВСК', 'НОВЫЙ УРЕНГОЙ', 'СТЕРЛИТАМАК', 'МАГНИТОГОРСК', 'ОРЕНБУРГ', 'СУРГУТ',
                   'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК'],
           'ПФО': ['ИЖЕВСК', 'ПЕНЗА', 'УЛЬЯНОВСК', 'ЧЕБОКСАРЫ', 'КИРОВ', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ', 'САМАРА',
                   'САРАТОВ', 'ТОЛЬЯТТИ'],
           'ЮФО': ['НОВОРОССИЙСК', 'СИМФЕРОПОЛЬ', 'ПЯТИГОРСК', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
                   'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ'],
           'СФО': ['БАРНАУЛ', 'НОВОКУЗНЕЦК', 'ТОМСК', 'УЛАН-УДЭ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
                   'КЕМЕРОВО'],
           'ДВФО': ['ВЛАДИВОСТОК', 'ХАБАРОВСК']}

city_dict = ['САНКТ-ПЕТЕРБУРГ', 'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ',
             'САМАРА', 'САРАТОВ', 'ТОЛЬЯТТИ', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
             'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
             'КЕМЕРОВО', 'ВЛАДИВОСТОК', 'ХАБАРОВСК']

def ret(cell):  # столбец и ячейку передаю, возрат - округ
	for i in dict_fo.keys():
		if str(cell).upper() in dict_fo[i]:
			return i
	else:
		return 'ЦФО'

def ower_city(row):
	if str(row[0]).upper() in city_dict:
		return 1
	elif row[1][:25] == 'Россия Московская область':
		return 1
	else:
		return 0

df.dropna(subset=['Общая стоимость со скидкой'], how='any', axis=0, inplace=True)

df['ФО'] = df['Заказ.Клиент.Подразделение.Адрес.Город'].apply(ret)
df['ФО'] = df['ФО'].astype('category')

cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)

df['Общая стоимость со скидкой'] = df['Общая стоимость со скидкой'].fillna(0)

df['Дата Cоздания'] = df['Дата Cоздания'].dt.strftime('%Y-%m-%d')

############ своя география
df['откуда'] = df.loc[:, ['Отправитель.Адрес.Город', 'Отправитель.Адрес']].apply(ower_city, axis=1)
df['куда'] = df.loc[:, ['Получатель.Адрес.Город', 'Получатель.Адрес']].apply(ower_city, axis=1)
# df.dropna(subset=['откуда','куда'], how='any', axis=0, inplace=True)

mask1 = df['откуда'] == 0
mask2 = df['куда'] == 0
df_agent = df[(mask1 | mask2)]
df_our = df[(~mask1 & ~mask2)]

print(round(df_our.shape[0]*100/df.shape[0], 1),'%')
#############

# df.rename(columns={'Дата Cоздания': 'дата',
#                    'Номер отправления': 'шт', 'Общая стоимость со скидкой': 'деньги', 'Расчетный вес': 'вес',
#                    'Отправитель.Адрес.Город': 'из', 'Получатель.Адрес.Город': 'в',
#                    'Заказ.Клиент.Не применять топливную надбавку': 'тн',
#                    'Заказ.Клиент.Подразделение.Адрес.Город': 'чей'}, inplace=True)

# df = df.loc[:, ['ФО', 'дата', 'шт', 'Клиент', 'деньги', 'вес', 'из', 'в', 'тн', 'чей']]

# mask1 = df['из'] == 'Москва'
# mask2 = df['в'] == 'Москва'
# df_dep = df[mask1]
# df_arr = df[mask2]
#
# df_inner = df[(mask1 & mask2)]
# df_or = df[(mask1 | mask2)]
#
# print(round(df_inner['деньги'].sum(),0))
# print(round(df_or['деньги'].sum(),0))
writer = pd.ExcelWriter('data/география.xlsx', engine='xlsxwriter')
df_our.to_excel(writer, sheet_name='итоги', index=True, header=True)
df_agent.to_excel(writer, sheet_name='агенты', index=True, header=True)
workbook = writer.book
writer.save()