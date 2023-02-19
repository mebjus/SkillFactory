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


def dw(row):
	if row == 0: d = 'понедельник'
	elif row == 1: d = 'вторник'
	elif row == 2: d = 'среда'
	elif row == 3: d = 'четверг'
	elif row == 4: d = 'пятница'
	elif row == 5: d = 'суббота'
	elif row == 6: d = 'воскресенье'
	else: d = np.NAN
	return d


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


df['ФО'] = df['Клиент.Подразделение.Адрес.Город'].apply(ret)
df['ФО'] = df['ФО'].astype('category')

cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)

df['week'] = df['Фактическая дата оплаты'].dt.dayofweek
df['week'] = df['week'].apply(dw)

# df['Дата счета'] = df['Дата счета'].dt.to_period('M')

####### фильтры

# df = df[df['Дата счета'] == '2021-09']
# df = df[df['ФО'] == 'СЗФО']
df = df[df['Статус оплаты счета'] == 'Оплачен']
df = df[df['Клиент'] == 'ООО "ЭСТЕ ЛАУДЕР КОМПАНИЗ"']

df['delay'] = df['Фактическая дата оплаты'] - df['Оплатить до']
df['delay']  = df['delay'].dt.components.days
df['delay'] = df['delay'].fillna(0)


df['delay'] = df['delay'].apply(lambda x: 0 if x < 0 else x)


print(df.info())

start = df.shape[0]
print(df['week'].value_counts(normalize=True)*100)
df1 = df.copy()

df = df[df['delay'] > 0]
finish = df.shape[0]/start
print('просроченных=', round(finish*100, 2), '%')

# llst = ['Статус счета','Статус оплаты счета','week','Сумма с учетом НДС', 'Клиент.Подразделение.Адрес.Город', 'Дата счета']
# df.drop(columns=llst, axis=1, inplace=True)
# df1.drop(columns=llst, axis=1, inplace=True)
#
# df = df.sort_values(by=['Клиент', 'Оплатить до'], ascending=[True, True])
# df1 = df1.sort_values(by=['Клиент', 'Оплатить до'], ascending=[True, True])
#

with pd.ExcelWriter('data/billing.xlsx') as writer:
	df1.to_excel(writer, sheet_name='оплаты')
