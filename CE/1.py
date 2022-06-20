import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel('data/DispatchesCount_online.xlsx', header=1)
df_m = pd.read_excel('data/temp/day_of_month.xlsx')

df = df.drop('Unnamed: 7', axis=1)
df = df.fillna(0)

dict_fo = {'СЗФО': ['Великий Новгород', 'Мурманск', 'Петрозаводск', 'Сыктывкар', 'Санкт-Петербург', 'Архангельск',
                    'Калининград'],
           'УФО': ['Курган', 'Нижневартовск', 'Новый Уренгой', 'Стерлитамак', 'Магнитогорск', 'Оренбург', 'Сургут',
                   'Екатеринбург', 'Пермь', 'Тюмень', 'Уфа', 'Челябинск'],
           'ПФО': ['Ижевск', 'Пенза', 'Ульяновск', 'Чебоксары', 'Киров', 'Нижний Новгород', 'Казань', 'Самара',
                   'Саратов', 'Тольятти'],
           'ЮФО': ['Новороссийск', 'Симферополь', 'Пятигорск', 'Ростов-на-Дону', 'Волгоград', 'Воронеж', 'Краснодар', 'Ставрополь',
                   'Астрахань', 'Сочи'],
           'СФО': ['Барнаул', 'Новокузнецк', 'Томск', 'Улан-Удэ', 'Новосибирск', 'Красноярск', 'Омск', 'Иркутск',
                   'Кемерово'],
           'ДВФО': ['Владивосток', 'Хабаровск'],
           'Итого': ['Итого']
           }

for i in dict_fo.keys():  # преобразование словаря верхний регистр
    d = []
    for j in dict_fo[i]:
        c = j.upper()
        d.append(c)
    dict_fo[i] = d


def ret(cell):  # столбец и ячейку передаю, возрат - округ
    for i in dict_fo.keys():
        if str(cell).upper() in dict_fo[i]:
            return i
    else:
        return 'ЦФО'


df['ФО'] = df['Unnamed: 0'].apply(ret)

for i in range(1, df.shape[0], 4):
    df.rename(index={i: df.iloc[i]['ФО'] + ' шт',
                     i + 1: df.iloc[i]['ФО'] + ' без скидки',
                     i + 2: df.iloc[i]['ФО'] + ' деньги',
                     i + 3: df.iloc[i]['ФО'] + ' кг',
                     }, inplace=True)

for i in range(1, df.shape[0], 4):
    df.rename(index={i: df.iloc[i][0] + ' шт',
                     i + 1: df.iloc[i][0] + ' без скидки',
                     i + 2: df.iloc[i][0] + ' деньги',
                     i + 3: df.iloc[i][0] + ' кг',
                     }, inplace=True)


df = df.T

def todate(arg):
    if arg != 0 and arg != 'ЦФО':
        y = pd.to_datetime(arg)
        y = y.strftime('%Y-%m')
        return y
    else:
        return arg



df[0] = df[0].apply(todate)

low = df[df['Итого шт'] == 'Итого']
df = df.drop(low.index, axis=0)

df.rename(columns={0: 'Дата'}, inplace=True)

df = df.groupby(['Дата'], as_index=False).sum().round(0)
df = df.merge(df_m, how='left')

df.replace({'Дата': {'ЦФО': 'Округ'}}, inplace=True)
# df.replace({'Итого шт': {'ЦФО': np.NAN}}, inplace=True)
# df.replace({'Итого без скидки': {'ЦФО': np.NAN}}, inplace=True)
# df.replace({'Итого деньги': {'ЦФО': np.NAN}}, inplace=True)
# df.replace({'Итого кг': {'ЦФО': np.NAN}}, inplace=True)

df['Итого, деньги на рд'] = (df['Итого деньги'] / df['count']).round(0)
df['Итого, кг на рд'] = (df['Итого кг'] / df['count']).round(0)

# fig = plt.figure(figsize=(15, 5))
#
# plt.subplot(121)
# sns.lineplot(x = 'Дата', y='Итого, кг на рд', data=df, lw=5)
# plt.xticks(rotation=45)
#
# plt.subplot(122)
# sns.lineplot(x = 'Дата', y='Итого, деньги на рд', data=df, lw=5)
# plt.xticks(rotation=45)
# plt.show()

# df=df.T
# df1 = df.pivot_table(index='Дата', columns='Округ', values='ЦФО деньги', aggfunc='sum')
# df1=df.groupby(by=['ЦФО кг'], axis=0).sum()
# print(df['СФО деньги'])
# df1=df['СФО деньги'].sum(axis=1)

# print(df1)

df.to_excel('./summary.xlsx', sheet_name='итоги', index=False)