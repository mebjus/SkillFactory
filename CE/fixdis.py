import numpy as np
import pandas as pd
import os
from pandas.api.types import CategoricalDtype


df = pd.DataFrame
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name : os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.0F}'.format

for file in fullpaths :
	if df.empty :
		if file.find('.xls') != -1 :
			df = pd.read_excel(file, header=2, sheet_name=None)
			df = pd.concat(df, axis=0).reset_index(drop=True)
	else :
		if file.find('.xls') != -1 :
			df1 = pd.read_excel(file, header=2, sheet_name=None)
			df1 = pd.concat(df1, axis=0).reset_index(drop=True)
			df = pd.concat([df, df1], axis=0)

df['Дата Cоздания'] = df['Дата Cоздания'].dt.strftime('%Y-%m')

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


def ret(cell) :  # столбец и ячейку передаю, возрат - округ
	for i in dict_fo.keys() :
		if str(cell).upper() in dict_fo[i] :
			return i
	else :
		return 'ЦФО'

df['ФО'] = df['Заказ.Клиент.Подразделение.Адрес.Город'].apply(ret)
df['ФО'] = df['ФО'].astype('category')

## установить порядок в списке ФО
cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
df['ФО'] = df['ФО'].astype(cat_type)

df.rename(columns={'Заказ.Клиент.Клиентский номер':'кн',
					'Дата Cоздания':'дата',
					'Номер отправления':'шт',
					'Общая стоимость со скидкой':'деньги',
					'Расчетный вес':'вес',
					'Заказ.Клиент.Подразделение.Адрес.Город':'город',
					'Заказ.Клиент.Не применять топливную надбавку':'тн'}, inplace=True)
df['тн'] = df['тн'].fillna(0)

# отбрасываем все нулевки, консолидированные сборы, дешевые доборы

# df = df[df['деньги'] > 50]
# df = df[df['вес'] <= 0.5]
# df = df[df['ФО'] == 'ЦФО']
# df = df[df['Вид доставки'] == 'Местная']

def disk(row):
	row['деньги'] = int(round(row['деньги'], 0))
	k = 0
	if (row['город'] == 'Москва') or (row['город'] == 'Санкт-Петербург'): k = 0
	else: k = 15
	if row['деньги'] in range(0, 10000) : return 5+k
	if row['деньги'] in range(10001, 20000) : return 6+k
	if row['деньги'] in range(20001, 30000) : return 7+k
	if row['деньги'] in range(30001, 40000) : return 9+k
	if row['деньги'] in range(40001, 50000) : return 10+k
	if row['деньги'] in range(50001, 60000) : return 11+k
	if row['деньги'] in range(60001, 70000) : return 12+k
	if row['деньги'] in range(70001, 80000) : return 13+k
	if row['деньги'] in range(80001, 90000) : return 15+k
	if row['деньги'] in range(90001, 100000) : return 16+k
	if row['деньги'] in range(100001, 110000) : return 17+k
	if row['деньги'] in range(110001, 120000) : return 18+k
	if row['деньги'] in range(120001, 130000) : return 19+k
	if row['деньги'] in range(130001, 140000) : return 21+k
	if row['деньги'] in range(140001, 150000) : return 22+k
	if row['деньги'] in range(150001, 160000) : return 23+k
	if row['деньги'] in range(160001, 170000) : return 24+k
	if row['деньги'] in range(170001, 180000) : return 25+k
	if row['деньги'] in range(180001, 190000) : return 27+k
	if row['деньги'] in range(190001, 200000) : return 29+k
	if row['деньги'] in range(200001, 210000) : return 30+k
	if row['деньги'] in range(210001, 220000) : return 33+k
	if row['деньги'] in range(220001, 230000) : return 34+k
	if row['деньги'] in range(230001, 240000) : return 35+k
	if row['деньги'] in range(240001, 250000) : return 36+k
	if row['деньги'] in range(250001, 260000) : return 37+k
	if row['деньги'] in range(260001, 270000) : return 39+k
	if row['деньги'] in range(270001, 280000) : return 40+k
	if row['деньги'] in range(280001, 290000) : return 41+k
	if row['деньги'] in range(290001, 300000) : return 42+k
	if row['деньги'] in range(310001, 320000) : return 43+k
	if row['деньги'] in range(320001, 330000) : return 45+k
	if row['деньги'] in range(330001, 340000) : return 46+k
	if row['деньги'] in range(340001, 350000) : return 47+k
	if row['деньги'] in range(350001, 360000) : return 48+k
	if row['деньги'] in range(360001, 370000) : return 49+k
	if row['деньги'] in range(370001, 380000) : return 51
	if row['деньги'] in range(380001, 390000) : return 52
	if row['деньги'] in range(390001, 400000) : return 53
	if row['деньги'] in range(400001, 410000) : return 54
	if row['деньги'] in range(410001, 420000) : return 55
	if row['деньги'] in range(420001, 430000) : return 57
	if row['деньги'] in range(430001, 440000) : return 58
	if row['деньги'] in range(440001, 450000) : return 59
	if row['деньги'] in range(450001, 460000) : return 60
	if row['деньги'] in range(460001, 470000) : return 61
	else:
		return 'перс'


df_pivot = df.pivot_table(index=['ФО', 'дата', 'город', 'кн', 'тн'], values=['деньги'], aggfunc={'деньги':sum})
df_pivot = df_pivot[df_pivot['деньги'] > 0]
df_pivot = df_pivot.reindex(df_pivot.sort_values(by=['ФО', 'дата', 'деньги'], ascending=[True, True, False]).index).reset_index()

### данные по клиентам (период счета, отсрочка и тп)

dirname = 'data/юл.xls'
df_tmp1 = pd.read_excel(dirname, sheet_name='Sheet1')
df_tmp2 = pd.read_excel(dirname, sheet_name='Sheet2')
df_tmp = pd.concat([df_tmp1, df_tmp2], axis=0)
df_tmp.reset_index()
df_pivot = df_pivot.merge(df_tmp, how='left', on='кн')
df_pivot.rename(columns={'Отсрочка платежа (дней)': 'отсрочка',
				   '(Юридическое лицо).Период формирования счетов': 'период'}, inplace=True)
df_pivot['тн'] = df_pivot['тн'].apply(lambda x: 'нет' if x == 1 else 'да')

###  подставляем глобальную скидку по клиентам
dirname = 'data/totaldis.xlsx'
df_tmp = pd.read_excel(dirname)
df_tmp.reset_index()
df_pivot = df_pivot.merge(df_tmp, how='left', on='Клиент')

df_pivot['dis'] = round(df_pivot['dis'] * -100, 0)
df_pivot['recomend'] = df_pivot.loc[:, ['деньги', 'город']].apply(disk, axis=1)
df_pivot = df_pivot.loc[:, ['ФО', 'дата', 'город', 'кн', 'Клиент', 'деньги', 'dis', 'recomend', 'тн', 'отсрочка', 'период']]


#### добавляем тех, у кого есть хотя бы один закрепленный тариф
dirname = 'data/фикс_скидка.xls'
df_tmp = pd.read_excel(dirname)
df_tmp.reset_index()
df_pivot = df_pivot.merge(df_tmp, how='left', on='кн')

# #### добавляем тех, у кого есть закрепленный прайс лист
dirname = 'data/hold_price.xls'
df_tmp = pd.read_excel(dirname)
df_tmp.reset_index()
df_pivot = df_pivot.merge(df_tmp, how='left', on='кн')

#####


dupl = list(df_pivot.columns)
df_dupl = df_pivot[df_pivot.duplicated(subset=dupl)]
df_dupl.index.nunique()
df_pivot = df_pivot.drop_duplicates(subset=dupl)

writer = pd.ExcelWriter('clients.xlsx', engine='xlsxwriter')
df_pivot.to_excel(writer, sheet_name='итоги', startrow=1, index=False, header=False)

workbook = writer.book
worksheet = writer.sheets['итоги']

format = workbook.add_format({'border' : 1, 'bg_color' : '#E8FBE1', 'num_format' : '#,##0'})
worksheet.set_column('A:E', 10, format)
worksheet.set_column('E:F', 65, format)
worksheet.set_column('F:P', 10, format)
header_format = workbook.add_format(
	{'bold' :          True, 'text_wrap' : True, 'valign' : 'vcenter', 'fg_color' : '#D7E4BC',
		'align' :      'center_across', 'num_format' : '#,##0', 'border' : 1})

for col_num, value in enumerate(df_pivot.columns.values) :
	worksheet.write(0, col_num, value, header_format)

writer.save()
