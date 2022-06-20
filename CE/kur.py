import numpy as np
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
df_m.reset_index()
mounth = {}

for i in df_m.index:
	mounth[df_m.iloc[i]['Дата']] = df_m.iloc[i]['р.д.']

df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('D')


def moscow(row):
	if row != 'nan' and str(row)[:3] == '770':
		return row
	else:
		return np.NAN


df['Прием курьером.Пакеты доставки.Курьер.Номер курьера'] = df[
	'Прием курьером.Пакеты доставки.Курьер.Номер курьера'].astype(str).apply(moscow)
df['Доставка курьером.Пакеты доставки.Курьер.Номер курьера'] = df[
	'Доставка курьером.Пакеты доставки.Курьер.Номер курьера'].astype(str).apply(moscow)


df1 = df.copy()

df = df.groupby('Дата Cоздания')['Прием курьером.Пакеты доставки.Курьер.Номер курьера'].agg(set1=lambda x: set(x))
df1 = df1.groupby('Дата Cоздания')['Доставка курьером.Пакеты доставки.Курьер.Номер курьера'].agg(set1=lambda x: set(x))

df['set2'] = df1['set1']

def discarded(row):
	if np.NAN in row: return row.discard(np.NAN)


df['set_all'] = df.apply(lambda x: x.set1.union(x.set2), axis=1)
df['количество'] = df['set_all'].apply(discarded)
df['количество'] = df['set_all'].apply(lambda x: len(x))
df = df.drop(columns=['set1', 'set2'])
df = df.reset_index()
print(df.info())
######

writer = pd.ExcelWriter('курьеры.xlsx', engine='xlsxwriter')
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
