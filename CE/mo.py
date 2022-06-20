import pandas as pd
import os

df = pd.DataFrame()
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.0F}'.format

for file in fullpaths:
    df1 = pd.read_excel(file, header=2, sheet_name=None)
    df1 = pd.concat(df1, axis=0).reset_index(drop=True)
    df = pd.concat([df, df1], axis=0)

dirname = 'data/day_of_month.xlsx'
df_m = pd.read_excel(dirname)

df_m['Дата'] = df_m['Дата'].dt.to_period('M')
df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('M')

df = df[df['Вид доставки'] == 'Международная']

df = df.groupby('Дата Cоздания')['Общая стоимость со скидкой'].sum()
df = df.reset_index()


df = df.merge(df_m, left_on='Дата Cоздания', right_on='Дата', how='left')
df['на рд']  = df['Общая стоимость со скидкой'] / df['р.д.']

df = df.loc[:, ['Дата Cоздания', 'на рд']]
print(df)
