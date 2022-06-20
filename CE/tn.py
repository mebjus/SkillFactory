import numpy as np
import pandas as pd
import os
from pandas.api.types import CategoricalDtype

df = pd.DataFrame
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.0F}'.format

for file in fullpaths:
	if df.empty:
		if file.find('.xls') != -1:
			df = pd.read_excel(file, header=2, sheet_name=None)
			df = pd.concat(df, axis=0).reset_index(drop=True)
	else:
		if file.find('.xls') != -1:
			df1 = pd.read_excel(file, header=2, sheet_name=None)
			df1 = pd.concat(df1, axis=0).reset_index(drop=True)
			df = pd.concat([df, df1], axis=0)

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
           'ДВФО': ['ВЛАДИВОСТОК', 'ХАБАРОВСК']
           }


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




####### фильтры

# df = df[df['Дата счета'] == '2021-09']
# df = df[df['Клиент'] == 'ООО "Эмарсис"']

#### 0 - не применять
df['Заказ.Клиент.Не применять топливную надбавку'] = df['Заказ.Клиент.Не применять топливную надбавку'].fillna(0)

print('с ТН, деньги:{:,.0F}'.format(
	df[df['Заказ.Клиент.Не применять топливную надбавку'] == 0]['Общая стоимость со скидкой'].sum()))  # c ТН
print('без ТН, деньги:{:,.0F}'.format(
	df[df['Заказ.Клиент.Не применять топливную надбавку'] == 1]['Общая стоимость со скидкой'].sum()))  # без ТН
# print(df['Общая стоимость со скидкой'].sum())



d = df['Заказ.Клиент.Не применять топливную надбавку'].value_counts(normalize=True)
print(d*100)

r = df[df['Заказ.Клиент.Не применять топливную надбавку'] == 0]
d = (r['Общая стоимость со скидкой'] - (r['Общая стоимость со скидкой'] / 122) * 100)
print('Денег от ТН:{:,.0F}'.format(d.sum()))

start = df.shape[0]

c1 = len(set(df['Клиент']))
df = df[df['Заказ.Клиент.Не применять топливную надбавку'] == 0]
c2 = len(set(df['Клиент']))
finish = df.shape[0] / start

print('Заказов с ТН от общего количества: ', round(finish * 100, 2), '%')
print('Клиентов с ТН от общего количества: ', round(c2/c1 * 100, 2), '%')

df['sizetn'] = (df['Общая стоимость со скидкой'] - (df['Общая стоимость со скидкой'] / 122) * 100)
df = df.reset_index()


lst = ['Заказ.Дата и время доставки', 'Получатель.Дата получения отправления получателем',
       'Отправитель.Дата приема у отправителя', 'Дата dead-line приема отправления', 'Получатель.Адрес',
       'Получатель.Адрес.Город']
df.drop(columns=lst, axis=1, inplace=True)


writer = pd.ExcelWriter('tn.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='итоги', startrow=1, index=False, header=False)

workbook = writer.book
worksheet = writer.sheets['итоги']

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

writer.save()
