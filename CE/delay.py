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
	df1 = pd.read_excel(file, header=2, sheet_name=None)
	df1 = pd.concat(df1, axis=0).reset_index(drop=True)
	df = pd.concat([df, df1], axis=0)

# df = df.dropna(how='any', axis=0)

dirname = 'data/day_of_month.xlsx'
df_m = pd.read_excel(dirname)
df_m.reset_index()
mounth = {}

for i in df_m.index:
	mounth[df_m.iloc[i]['Дата']] = df_m.iloc[i]['р.д.']

df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('D')

##### фильтр на дату

# df = df[df['Дата Cоздания'] == '2021-11']

######

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

city_dict = ['САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК',
             'КАЛИНИНГРАД', 'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ',
             'САМАРА',
             'САРАТОВ', 'ТОЛЬЯТТИ', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
             'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
             'КЕМЕРОВО', 'ВЛАДИВОСТОК', 'ХАБАРОВСК', 'МОСКВА']


def ret(cell):  # столбец и ячейку передаю, возрат - округ
	for i in dict_fo.keys():
		if str(cell).upper() in dict_fo[i]:
			return i
	else:
		return 'ЦФО'


df['ФО'] = df['Заказ.Клиент.Подразделение.Адрес.Город'].apply(ret)
df['ФО'] = df['ФО'].astype('category')


def ower_city(row):
	if str(row).upper() not in city_dict:
		return np.NAN
	else:
		return row


#### фильтр на свою географю

# df['Отправитель.Адрес.Город'] = df['Отправитель.Адрес.Город'].apply(ower_city)
# df['Получатель.Адрес.Город'] = df['Получатель.Адрес.Город'].apply(ower_city)
# df = df.dropna(how='any', axis=0)


# установить порядок в списке ФО
cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)

df['Группа вес'] = pd.cut(df['Расчетный вес'], bins=[0, 1, 5, 30, 100, 1000000],
                          labels=['0-1', '1-5', '5-30', '30-100', '100+'], right=False)

df['Группа вес'] = df['Группа вес'].astype('category')

## установить порядок по весам
cat_type = CategoricalDtype(categories=['0-1', '1-5', '5-30', '30-100', '100+'], ordered=True)
df['Группа вес'] = df['Группа вес'].astype(cat_type)

df.rename(columns={'Дата Cоздания':     'дата',
                   'Номер отправления': 'шт', 'Общая стоимость со скидкой': 'деньги', 'Расчетный вес': 'вес'},
          inplace=True)


def mod(arg):
	if arg.find('ЭКСПРЕСС') != -1:
		return 'ЭКСПРЕСС'
	elif arg.find('ПРАЙМ') != -1:
		return 'ПРАЙМ'
	elif arg.find('ОПТИМА') != -1:
		return 'ОПТИМА'
	else:
		return 'ПРОЧИЕ'


df['Режим доставки'] = df['Режим доставки'].apply(mod)
df['Режим доставки'] = df['Режим доставки'].astype('category')

# отбрасываем все условия:

df = df[df['деньги'] > 50]
# df = df[df['вес'] <= 0.250]
# df = df[df['ФО'] == 'ЦФО']
# df = df[df['Режим доставки'] == 'ЭКСПРЕСС']
# df = df[df['Вид доставки'] == 'Местная']
df = df[df['Клиент'] == 'ООО "РОКВУЛ"']

#

num_start = df.shape[0]
print(df.shape[0])

#### если положительное значение - просрочка

df['delta_delivery'] = df['Получатель.Дата получения отправления получателем'] - df['Заказ.Дата и время доставки']
df['delta_get'] = df['Отправитель.Дата приема у отправителя'] - df['Дата dead-line приема отправления']
df['delta_delivery'] = df['delta_delivery'].dt.components.days
df['delta_get'] = df['delta_get'].dt.components.days

df1 = df[df['delta_delivery'] > 0]
df1 = df1.reset_index()
df1.drop(columns=['index', 'Режим доставки', 'Скидка', 'Вид доставки', 'Получатель.Адрес',
                  'Заказ.Клиент.Не применять топливную надбавку', 'Отправитель.Дата приема у отправителя',
                  'Дата dead-line приема отправления', 'Заказ.Дата и время доставки',
                  'Получатель.Дата получения отправления получателем', 'Группа вес', 'delta_get'], axis=1, inplace=True)
print(round((df1['delta_delivery'].shape[0] / num_start) * 100, 2), '%', 'нарушены сроки доставки')

df2 = df[df['delta_get'] > 0]
df2 = df2.reset_index()
df2.drop(columns=['index', 'Режим доставки', 'Скидка', 'Вид доставки', 'Получатель.Адрес',
                  'Заказ.Клиент.Не применять топливную надбавку', 'Отправитель.Дата приема у отправителя',
                  'Дата dead-line приема отправления', 'Заказ.Дата и время доставки',
                  'Получатель.Дата получения отправления получателем', 'Группа вес', 'delta_delivery'],
         axis=1,
         inplace=True)
print(round((df2['delta_get'].shape[0] / num_start) * 100, 2), '%', 'нарушены сроки сбора')

#######

writer = pd.ExcelWriter('delay.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name='не доставки', startrow=1, index=False, header=False)
df2.to_excel(writer, sheet_name='не сборы', startrow=1, index=False, header=False)

workbook = writer.book
worksheet = writer.sheets['не доставки']
worksheet2 = writer.sheets['не сборы']

header_format = workbook.add_format({
	'bold':       True,
	'text_wrap':  True,
	'valign':     'vcenter',
	'fg_color':   '#D7E4BC',
	'align':      'center_across',
	'num_format': '#,##0',
	'border':     1})

for col_num, value in enumerate(df1.columns.values):
	worksheet.write(0, col_num, value, header_format)
for col_num, value in enumerate(df2.columns.values):
	worksheet2.write(0, col_num, value, header_format)

writer.save()
