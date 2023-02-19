import pandas as pd
import os

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

day_now = pd.to_datetime('now')
df['Дата Cоздания'] = df['Дата Cоздания'].dt.to_period('D')

######
df.dropna(subset=['Получатель.Дата получения отправления получателем', 'Отправитель.Дата приема у отправителя'],
          how='any', axis=0, inplace=True)

# df['Отправитель.Дата приема у отправителя'].fillna('nan', inplace=True)
#
# def onnow(row):
#     if row[0] != 'nan' and row[1] < day_now: return day_now
#
#
# df['Получатель.Дата получения отправления получателем'] = df.loc[:, ['Отправитель.Дата приема у отправителя',
#                                                                      'Заказ.Дата и время доставки']].apply(onnow,
#                                                                                                            axis=1)
# df1 = df.copy()
# for i in df.index:
# 	if df.iloc[i,15] == 'nan':
# 		df1.drop(labels=i, axis=0, inplace=True)
# df = df1.copy()

df = df[df['Общая стоимость со скидкой'] > 0]
# df = df[df['Клиент'] == 'ООО "Петролеум Трейдинг"']

#### если положительное значение - просрочка
df['delta_delivery'] = (df['Получатель.Дата получения отправления получателем'] - df[
	'Заказ.Дата и время доставки']).dt.components.days
df['delta_get'] = (
		df['Отправитель.Дата приема у отправителя'] - df['Дата dead-line приема отправления']).dt.components.days

df1 = df[df['delta_delivery'] > 0]
print(round((df1.shape[0] / df.shape[0]) * 100, 2), '%', 'нарушены сроки доставки')

df1 = df[df['delta_get'] > 0]
print(round((df1.shape[0] / df.shape[0]) * 100, 2), '%', 'нарушены сроки сбора')

#######
with pd.ExcelWriter('data/delay.xlsx') as writer:
	df.to_excel(writer, sheet_name='нарушения')
