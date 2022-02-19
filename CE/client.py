import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pandas.api.types import CategoricalDtype
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

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

dirname = 'data/day_of_month.xlsx'
df_m = pd.read_excel(dirname)
df_m.reset_index()
mounth = {}

for i in df_m.index:
    mounth[df_m.iloc[i]['Дата']] = df_m.iloc[i]['р.д.']

dupl = list(df.columns)
df_dupl = df[df.duplicated(subset=dupl)]
df_dupl.index.nunique()
df = df.drop_duplicates(subset=dupl)

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


def ret(cell):  # столбец и ячейку передаю, возрат - округ
    for i in dict_fo.keys():
        if str(cell).upper() in dict_fo[i]:
            return i
    else:
        return 'ЦФО'


df['ФО'] = df['Заказ.Клиент.Подразделение.Адрес.Город'].apply(ret)
df['ФО'] = df['ФО'].astype('category')

## установить порядок в списке ФО
cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)

df['Группа вес'] = pd.cut(df['Расчетный вес'], bins=[0, 1, 5, 30, 100, 1000000],
                          labels=['0-1', '1-5', '5-30', '30-100', '100+'], right=False)

df['Группа вес'] = df['Группа вес'].astype('category')

## установить порядок в по весам
cat_type = CategoricalDtype(categories=['0-1', '1-5', '5-30', '30-100', '100+'], ordered=True)
df['Группа вес'] = df['Группа вес'].astype(cat_type)

df.rename(columns={'Дата Cоздания': 'дата',
                   'Номер отправления': 'шт', 'Общая стоимость со скидкой': 'деньги', 'Расчетный вес': 'вес'},
          inplace=True)

# отбрасываем все нулевки, консолидированные сборы, дешевые доборы

df = df[df['деньги'] > 50]
# df = df[df['вес'] <= 0.5]
# df = df[df['ФО'] == 'ЦФО']
# df = df[df['Вид доставки'] == 'Местная']


df_pivot = df.pivot_table(index=['дата', 'Клиент'], values=['деньги', 'шт', 'вес'],
                          aggfunc={'деньги': sum, 'шт': len, 'вес': sum})

df_pivot = df_pivot.reindex(df_pivot.sort_values(by=['дата', 'деньги'], ascending=[True, False]).index).reset_index()

df_pivot['р.д.'] = df_pivot['дата'].apply(lambda x: mounth[str(x)])
df_pivot['деньги р.д.'] = df_pivot['деньги'] / df_pivot['р.д.']

df_pivot['ср чек'] = df_pivot['деньги'] / df_pivot['шт']

# df_pivot = df_pivot[df_pivot['ср чек'] < 5000]     ### для отчета 0,25


# df_pivot['шт р.д.'] = df_pivot['шт'] / df_pivot['р.д.']
# df_pivot['вес р.д.'] = df_pivot['вес'] / df_pivot['р.д.']

# ##### сюда если конкретного клиента, но надо выборку за день делать
#
# name = 'Индивидуальный предприниматель Саванеев Вячеслав Владимирович'
# df_pivot = df_pivot[df_pivot['Клиент'] == name]

# сюда если всех клиентов, но надо выборку за месяц делать

# df_pivot = df_pivot.groupby('дата').agg(
#     {'Клиент': 'count', 'шт р.д.': 'sum', 'вес р.д.': 'sum', 'деньги': 'sum', 'деньги р.д.': 'sum'})
# df_pivot = df_pivot.reset_index()


######

yaxes = df_pivot.groupby('дата')['деньги'].sum().reset_index()
yaxes['дата'] = yaxes['дата'].astype('str')
print(yaxes['дата'])

######

fig, ax = plt.subplots(figsize=(8, 5))
plt.xticks(rotation=45)
g = sns.barplot(data=yaxes, x='дата', y='деньги', color='green')
ticks_loc = ax.get_yticks().tolist()
ax.yaxis.set_major_locator(ticker.FixedLocator(ticks_loc))
ylabels = ['{:,.0f}'.format(x) for x in g.get_yticks()]
g.set_yticklabels(ylabels)
plt.show()

#######

writer = pd.ExcelWriter('clients.xlsx', engine='xlsxwriter')
df_pivot.to_excel(writer, sheet_name='итоги', startrow=1, index=False, header=False)

workbook = writer.book
worksheet = writer.sheets['итоги']

format = workbook.add_format({'border': 1, 'bg_color': '#E8FBE1', 'num_format': '#,##0'})
worksheet.set_column('A:B', 10, format)
worksheet.set_column('B:C', 65, format)
worksheet.set_column('C:I', 15, format)

header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'vcenter',
    'fg_color': '#D7E4BC',
    'align': 'center_across',
    'num_format': '#,##0',
    'border': 1})

for col_num, value in enumerate(df_pivot.columns.values):
    worksheet.write(0, col_num, value, header_format)

writer.save()
