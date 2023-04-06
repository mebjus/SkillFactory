import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pandas.api.types import CategoricalDtype
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

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

dirname = 'data/day_of_month.xlsx'
df_m = pd.read_excel(dirname)
df_m.reset_index()
mounth = {}

df_m['Дата'] = df_m['Дата'].dt.strftime('%Y-%m')
for i in df_m.index :
	mounth[df_m.iloc[i]['Дата']] = df_m.iloc[i]['р.д.']

df['Дата Cоздания'] = pd.to_datetime(df['Дата Cоздания']).dt.strftime('%Y-%m')
df['р.д.'] = df['Дата Cоздания'].apply(lambda x: mounth[str(x)])

dict_fo = {'СЗФО' : ['ВЕЛИКИЙ НОВГОРОД', 'МУРМАНСК', 'ПЕТРОЗАВОДСК', 'СЫКТЫВКАР', 'САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК',
                     'КАЛИНИНГРАД'],
           'УФО' : ['КУРГАН', 'НИЖНЕВАРТОВСК', 'НОВЫЙ УРЕНГОЙ', 'СТЕРЛИТАМАК', 'МАГНИТОГОРСК', 'ОРЕНБУРГ', 'СУРГУТ',
                    'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК'],
           'ПФО' : ['ИЖЕВСК', 'ПЕНЗА', 'УЛЬЯНОВСК', 'ЧЕБОКСАРЫ', 'КИРОВ', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ', 'САМАРА', 'САРАТОВ',
                    'ТОЛЬЯТТИ'],
           'ЮФО' : ['НОВОРОССИЙСК', 'СИМФЕРОПОЛЬ', 'ПЯТИГОРСК', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
                    'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ'],
           'СФО' : ['БАРНАУЛ', 'НОВОКУЗНЕЦК', 'ТОМСК', 'УЛАН-УДЭ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК', 'КЕМЕРОВО'],
           'ДВФО' : ['ВЛАДИВОСТОК', 'ХАБАРОВСК']}

city_dict = ['САНКТ-ПЕТЕРБУРГ', 'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ',
             'САМАРА', 'САРАТОВ', 'ТОЛЬЯТТИ', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
             'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
             'КЕМЕРОВО', 'ВЛАДИВОСТОК', 'ХАБАРОВСК']


def ret(cell) :  # столбец и ячейку передаю, возрат - округ
	for i in dict_fo.keys() :
		if str(cell).upper() in dict_fo[i] :
			return i
	else :
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

## установить порядок в списке ФО
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


df['Режим доставки'] = df['Режим доставки'].apply(mod)
df['Режим доставки'] = df['Режим доставки'].astype('category')

## установить порядок в по режиму
cat_type = CategoricalDtype(categories=['ПРАЙМ', 'ЭКСПРЕСС', 'ОПТИМА'], ordered=True)
df['Режим доставки'] = df['Режим доставки'].astype(cat_type)


## установить порядок в по весам
df['Группа вес'] = pd.cut(df['Расчетный вес'], bins=[0, 0.5, 1, 5, 30, 100, 1000000],
                          labels=['0-0.5', '0.5-1', '1-5', '5-30', '30-100', '100+'], right=False)
df['Группа вес'] = df['Группа вес'].astype('category')
cat_type = CategoricalDtype(categories=['0-0.5', '0.5-1', '1-5', '5-30', '30-100', '100+'], ordered=True)
df['Группа вес'] = df['Группа вес'].astype(cat_type)


df.rename(columns={'Дата Cоздания' : 'дата', 'Номер отправления' : 'шт', 'Общая стоимость со скидкой' : 'деньги',
                   'Расчетный вес':'вес'}, inplace=True)

############ своя география

# df = df[df['Вид доставки'] != 'Местная']

df['откуда'] = df.loc[:, ['Отправитель.Адрес.Город', 'Отправитель.Адрес']].apply(ower_city, axis=1)
df['куда'] = df.loc[:, ['Получатель.Адрес.Город', 'Получатель.Адрес']].apply(ower_city, axis=1)

mask1 = df['откуда'] == 0
mask2 = df['куда'] == 0

df_agent = df[(mask1 | mask2)]
df_our = df[(~mask1 & ~mask2)]

print(round(df_our.shape[0]*100/df.shape[0], 1),'%', ' - наша сеть')
#############


#### при копировании df выабрать география - наша или нет или общая

##### по группе веса
df_w = df.copy()
# df_w = df_our.copy()
# df_w = df_agent.copy()

df_w = df_w.loc[:, ['деньги', 'вес', 'ФО','Группа вес']]
df_w = df_w.groupby(['ФО','Группа вес']).agg([np.sum, len])

df_w['средний чек'] = df_w[('деньги', 'sum')]/df_w[('деньги', 'len')]
df_w['средний кг'] = df_w[('деньги', 'sum')]/df_w[('вес', 'sum')]
df_w.drop([('вес', 'len')], axis=1, inplace=True)
df_w.rename(columns={('len'):('кол-во')}, inplace=True)

##### по режимам доставки
df_r = df.copy()
# df_r = df_our.copy()
# df_r = df_agent.copy()

# fig, ax = plt.subplots(nrows=2, ncols=1)
# fig.suptitle('деньги / вес')
# ax[0].pie(df_diagr[('sum', 'деньги')], labels=df_diagr.index, autopct='%1.2f%%')
# ax[1].pie(df_diagr[('sum', 'вес')], labels=df_diagr.index, autopct='%1.2f%%')
# plt.show()

df_r = df_r.pivot_table(values=['деньги', 'вес'], index=['ФО','Режим доставки'], aggfunc=[np.sum, len])
df_r.drop([('len', 'вес')], axis=1, inplace=True)
df_r.rename(columns={('len'):('кол-во')}, inplace=True)

####################
with pd.ExcelWriter('data/итоги месяца.xlsx') as writer:
    df_w.to_excel(writer, sheet_name="итоги")
    df_r.to_excel(writer, sheet_name="режим доставки")
