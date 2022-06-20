import pandas as pd
import numpy as np
import os
import datetime

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

start = '2021-03-01'
finish = '2021-03-31'

rd = 22  # март 2022
### лимит на стоп 15 мин
t = datetime.datetime.strptime('00:15:00', "%H:%M:%S")
td2 = datetime.timedelta(hours=t.hour, minutes=t.minute, seconds=t.second)

df['Прием курьером.Пакеты доставки.Курьер.Номер курьера'] = df[
    'Прием курьером.Пакеты доставки.Курьер.Номер курьера'].astype(str)
df['Доставка курьером.Пакеты доставки.Курьер.Номер курьера'] = df[
    'Доставка курьером.Пакеты доставки.Курьер.Номер курьера'].astype(str)


def chek_ufa(row):
    if row[2:3] == '.':
        return '0' + row
    else:
        return row


df['Прием курьером.Пакеты доставки.Курьер.Номер курьера'] = df[
    'Прием курьером.Пакеты доставки.Курьер.Номер курьера'].apply(chek_ufa)
df['Доставка курьером.Пакеты доставки.Курьер.Номер курьера'] = df[
    'Доставка курьером.Пакеты доставки.Курьер.Номер курьера'].apply(chek_ufa)

df.rename(columns={'Номер отправления': 'шт', 'Общая стоимость со скидкой': 'деньги',
                   'Расчетный вес': 'вес', 'Отправитель.Адрес.Город': 'город откуда',
                   'Получатель.Адрес.Город': 'город куда',
                   'Получатель.Дата получения отправления получателем': 'дата получения',
                   'Прием курьером.Пакеты доставки.Курьер.Номер курьера': 'прием',
                   'Доставка курьером.Пакеты доставки.Курьер.Номер курьера': 'доставка'}, inplace=True)


# процедура проверки курьера из этого ли города отправление
def chek_city(row):
    if row['шт'][0:2] != row['прием'][0:2]:
        return 'nan'
    else:
        return row['прием']
df['прием'] = df.loc[:, ['шт', 'прием']].apply(chek_city, axis=1)

### сборы, консолидация
df_pick = df[df['Отправитель.Дата приема у отправителя'] >= start]
df_pick = df_pick[df_pick['Отправитель.Дата приема у отправителя'] <= finish]

df_pick.loc[:, 'tmp'] = df_pick.loc[:, 'Отправитель.Дата приема у отправителя'].dt.strftime('%Y-%m-%d')
df_pick = df_pick.sort_values(['tmp', 'Отправитель.Адрес', 'Отправитель.Дата приема у отправителя'])

df_pick['diff'] = df_pick['Отправитель.Дата приема у отправителя'].diff()
df_pick = df_pick.loc[:,
          ['tmp', 'шт', 'Клиент', 'город откуда', 'Отправитель.Адрес', 'Отправитель.Дата приема у отправителя', 'diff',
           'прием']]

df_pick['консолид'] = 0  # по умолчанию все сконсолидировано

for i in range(df_pick.shape[0]):
    if df_pick.iloc[i, 6] > td2 or df_pick.iloc[i, 6] < datetime.timedelta():
        df_pick.iloc[i, 8] = 1
    if df_pick.iloc[i - 1, 4] != df_pick.iloc[i, 4]:
        df_pick.iloc[i, 6] = datetime.timedelta()
        df_pick.iloc[i, 8] = 1

# df_pick.dropna(subset=['tmp'], axis=0, how='any', inplace=True)  ### незаполненные данные по сбору
df_pick_group = df_pick.groupby(['tmp', 'город откуда'])['консолид'].sum().reset_index()
df_pick_group = df_pick_group.reindex(df_pick_group.sort_values(by=['tmp', 'консолид'], ascending=[True, False]).index)

### доставка, консолидация
df_deliv = df[df['дата получения'] >= start]
df_deliv = df_deliv[df_deliv['дата получения'] <= finish]

df_deliv.loc[:, 'tmp'] = df_deliv.loc[:, 'дата получения'].dt.strftime('%Y-%m-%d')
df_deliv = df_deliv.sort_values(['tmp', 'Получатель.Адрес', 'дата получения'])

df_deliv['diff'] = df_deliv['дата получения'].diff()
df_deliv = df_deliv.loc[:,
           ['tmp', 'шт', 'Клиент', 'город куда', 'Получатель.Адрес', 'дата получения', 'diff', 'доставка']]

df_deliv['консолид'] = 0  # по умолчанию все сконсолидировано

for i in range(df_deliv.shape[0]):
    if df_deliv.iloc[i, 6] > td2 or df_deliv.iloc[i, 6] < datetime.timedelta():
        df_deliv.iloc[i, 8] = 1
    if df_deliv.iloc[i - 1, 4] != df_deliv.iloc[i, 4]:
        df_deliv.iloc[i, 6] = datetime.timedelta()
        df_deliv.iloc[i, 8] = 1

df_deliv_group = df_deliv.groupby(['tmp', 'город куда'])['консолид'].sum().reset_index()
df_deliv_group = df_deliv_group.reindex(
    df_deliv_group.sort_values(by=['tmp', 'консолид'], ascending=[True, False]).index)

###############
df_pick_group.rename(columns={'город откуда': 'город'}, inplace=True)
df_deliv_group.rename(columns={'город куда': 'город'}, inplace=True)
df_all = pd.concat([df_deliv_group, df_pick_group], axis=0)
df_all = df_all.groupby(['tmp', 'город'])['консолид'].sum().reset_index()
df_all = df_all.reindex(df_all.sort_values(by=['tmp', 'консолид'], ascending=[True, False]).index)

###  кол-во курьеров на сборе
df_curr = df_pick.pivot_table(index=['tmp', 'город откуда'], values=['прием'], aggfunc={'прием': set}).reset_index()
df_curr.loc[:, 'count'] = df_curr.loc[:, 'прием'].apply(lambda x: len(x) - 1 if 'nan' in x else len(x))
df_curr = df_curr.reindex(df_curr.sort_values(by=['tmp', 'count'], ascending=[True, False]).index)

###  кол-во курьеров на доставке
df_curr2 = df_deliv.pivot_table(index=['tmp', 'город куда'], values=['доставка'],
                                aggfunc={'доставка': set}).reset_index()
df_curr2.loc[:, 'count'] = df_curr2.loc[:, 'доставка'].apply(lambda x: len(x) - 1 if 'nan' in x else len(x))
df_curr2 = df_curr2.reindex(df_curr2.sort_values(by=['tmp', 'count'], ascending=[True, False]).index)

###  кол-во курьеров итого
df_curr.rename(columns={'город откуда': 'город', 'прием': 'set1'}, inplace=True)
df_curr2.rename(columns={'город куда': 'город', 'доставка': 'set2'}, inplace=True)
df_curr_all = df_curr.merge(df_curr2, on=['tmp', 'город'])
df_curr_all['множество'] = df_curr_all.apply(lambda x: x.set1.union(x.set2), axis=1)
df_curr_all = df_curr_all.loc[:, ['tmp', 'город', 'множество']]

# df_curr_all['count'] = df_curr_all['множество'].apply(lambda x: len(x))
df_curr_all.loc[:, 'count'] = df_curr_all.loc[:, 'множество'].apply(lambda x: len(x) - 1 if 'nan' in x else len(x))

####  выработка на курьера
df_all = df_all.merge(df_curr_all, on=['tmp', 'город'])
df_all['стопы'] = round(df_all['консолид'] / df_all['count'], 0)
df_all = df_all.loc[:, ['tmp', 'город', 'консолид', 'count', 'стопы']]
df_all.rename(columns={'tmp': 'дата', 'консолид': 'сбор+доставка', 'count': 'N курьеров'}, inplace=True)

#####
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
df_pick.to_excel(writer, sheet_name='сбор', startrow=0, index=False, header=True)  ## общие данные
df_pick_group.to_excel(writer, sheet_name='сбор_с консолид', startrow=0, index=False, header=True)
df_deliv.to_excel(writer, sheet_name='доставка', startrow=0, index=False, header=True)
df_deliv_group.to_excel(writer, sheet_name='доставка_с консолид', startrow=0, index=False, header=True)
df_all.to_excel(writer, sheet_name='стопы', startrow=0, index=False, header=True)
df_curr.to_excel(writer, sheet_name='курьеры сборы', startrow=0, index=False, header=True)
df_curr2.to_excel(writer, sheet_name='курьеры доставка', startrow=0, index=False, header=True)
df_curr_all.to_excel(writer, sheet_name='курьеры total', startrow=0, index=False, header=True)

workbook = writer.book
writer.save()
