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

df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('M')

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


########    выбор по своей географии

def ower_city(row):
    if str(row).upper() not in city_dict:
        return np.NAN
    else:
        return row

df['Отправитель.Адрес.Город'] = df['Отправитель.Адрес.Город'].apply(ower_city)
df['Получатель.Адрес.Город'] = df['Получатель.Адрес.Город'].apply(ower_city)
df = df.dropna(how='any', axis=0)


## установить порядок в списке ФО
cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)

df['Группа вес'] = pd.cut(df['Расчетный вес'], bins=[0, 0.1, 0.15, 0.2, 0.25, 0.5, 1, 5, 30, 100, 1000000],
                          labels=['0-0,1', '0,1-0,15', '0,15-0,2',
                                  '0.2-0,25', '0,25-0,5', '0,5-1',
                                  '1-5', '5-30','30-100', '100+'], right=False)

df['Группа вес'] = df['Группа вес'].astype('category')

## установить порядок в по весам
cat_type = CategoricalDtype(
    categories=['0-0,1', '0,1-0,15', '0,15-0,2', '0.2-0,25', '0,25-0,5', '0,5-1', '1-5', '5-30', '30-100', '100+'],
    ordered=True)
df['Группа вес'] = df['Группа вес'].astype(cat_type)

df.rename(columns={'Дата Cоздания': 'дата',
                   'Номер отправления': 'шт', 'Общая стоимость со скидкой': 'деньги', 'Расчетный вес': 'вес'},
          inplace=True)

# отбрасываем все нулевки, консолидированные сборы, дешевые доборы

df = df[df['деньги'] > 10]

# df = df[df['Заказ.Клиент.Подразделение.Адрес.Город'] == 'Москва']
# df = df[df['дата'] == '2021-10']

df_pivot = df.pivot_table(index=['Вид доставки', 'дата'], columns=['Группа вес'], values=['деньги', 'шт'],
                          aggfunc={'деньги': sum, 'шт': len})


df_pivot[('деньги', '0-0,1%')] = (df_pivot[('деньги', '0-0,1')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '0,1-0,15%')] = (df_pivot[('деньги', '0,1-0,15')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '0,15-0,2%')] = (df_pivot[('деньги', '0,15-0,2')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '0.2-0,25%')] = (df_pivot[('деньги', '0.2-0,25')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '0,25-0,5%')] = (df_pivot[('деньги', '0,25-0,5')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '0,5-1%')] = (df_pivot[('деньги', '0,5-1')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '1-5%')] = (df_pivot[('деньги', '1-5')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '5-30%')] = (df_pivot[('деньги', '5-30')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '30-100%')] = (df_pivot[('деньги', '30-100')] / df_pivot[('деньги',)].sum(axis=1))
df_pivot[('деньги', '100+%')] = (df_pivot[('деньги', '100+')] / df_pivot[('деньги',)].sum(axis=1))


######   запись в файл

writer = pd.ExcelWriter('0.25.xlsx', engine='xlsxwriter')
df_pivot.to_excel(writer, sheet_name='итоги', startrow=0)
workbook = writer.book
worksheet = writer.sheets['итоги']

writer.save()
