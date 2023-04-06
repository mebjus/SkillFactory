import pandas as pd
import os
from pandas.api.types import CategoricalDtype

pd.options.display.float_format = '{:,.0F}'.format

class Readdf:
	__df = None
	dict_fo = {
		'СЗФО': ['ВЕЛИКИЙ НОВГОРОД', 'МУРМАНСК', 'ПЕТРОЗАВОДСК', 'СЫКТЫВКАР', 'САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК',
			'КАЛИНИНГРАД'],
		'УФО': ['КУРГАН', 'НИЖНЕВАРТОВСК', 'НОВЫЙ УРЕНГОЙ', 'СТЕРЛИТАМАК', 'МАГНИТОГОРСК', 'ОРЕНБУРГ', 'СУРГУТ',
			'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА', 'ЧЕЛЯБИНСК'],
		'ПФО': ['ИЖЕВСК', 'ПЕНЗА', 'УЛЬЯНОВСК', 'ЧЕБОКСАРЫ', 'КИРОВ', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ', 'САМАРА', 'САРАТОВ',
			'ТОЛЬЯТТИ'],
		'ЮФО': ['НОВОРОССИЙСК', 'СИМФЕРОПОЛЬ', 'ПЯТИГОРСК', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД', 'ВОРОНЕЖ', 'КРАСНОДАР',
			'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ'],
		'СФО': ['БАРНАУЛ', 'НОВОКУЗНЕЦК', 'ТОМСК', 'УЛАН-УДЭ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
			'КЕМЕРОВО'], 'ДВФО': ['ВЛАДИВОСТОК', 'ХАБАРОВСК']}
	
	city_dict = ['САНКТ-ПЕТЕРБУРГ', 'АРХАНГЕЛЬСК', 'КАЛИНИНГРАД', 'ЕКАТЕРИНБУРГ', 'ПЕРМЬ', 'ТЮМЕНЬ', 'УФА',
		'ЧЕЛЯБИНСК', 'НИЖНИЙ НОВГОРОД', 'КАЗАНЬ', 'САМАРА', 'САРАТОВ', 'ТОЛЬЯТТИ', 'РОСТОВ-НА-ДОНУ', 'ВОЛГОГРАД',
		'ВОРОНЕЖ', 'КРАСНОДАР', 'СТАВРОПОЛЬ', 'АСТРАХАНЬ', 'СОЧИ', 'НОВОСИБИРСК', 'КРАСНОЯРСК', 'ОМСК', 'ИРКУТСК',
		'КЕМЕРОВО', 'ВЛАДИВОСТОК', 'ХАБАРОВСК', 'МОСКВА']
		
	def __new__(cls, *args, **kwargs):
		if cls.__df is None:
			cls.__df = super().__new__(cls)
		return cls.__df
	
	def __init__(self, dirname = 'data/kis/'):
		self.df = pd.DataFrame()
		self.dirname = dirname
		
		dirfiles = os.listdir(self.dirname)
		fullpaths = map(lambda name: os.path.join(self.dirname, name), dirfiles)
		
		for file in fullpaths:
			if file == 'data/kis/.DS_Store': os.remove('data/kis/.DS_Store')
			self.df1 = pd.read_excel(file, header=2, sheet_name=None)
			self.df1 = pd.concat(self.df1, axis=0).reset_index(drop=True)
			self.df = pd.concat([self.df, self.df1], axis=0)
	
	def __getitem__(self, item):
		return self.df.iloc[item, :]
	
	@classmethod
	def ret(cls, cell):  # столбец и ячейку передаю, возрат - округ
		for _ in cls.dict_fo.keys():
			if str(cell).upper() in cls.dict_fo[_]:
				return _
		else:
			return 'ЦФО'
	
	@classmethod
	def ower_city(cls, row):
		if str(row[0]).upper() in cls.city_dict:
			return 1
		elif row[1][:25] == 'Россия Московская область':
			return 1
		else:
			return 0
			
	def fo(self, own=False):
		self.df['ФО'] = self.df['Заказ.Клиент.Подразделение.Адрес.Город'].apply(Readdf.ret)
		self.df['ФО'] = self.df['ФО'].astype('category')
		cat_type = CategoricalDtype(categories=['ЦФО', 'СЗФО', 'ПФО', 'ЮФО', 'УФО', 'СФО', 'ДВФО'], ordered=True)
		self.df['ФО'] = self.df['ФО'].astype(cat_type)

		if own:
			self.df['откуда'] = self.df.loc[:, ['Отправитель.Адрес.Город', 'Отправитель.Адрес']].apply(Readdf.ower_city, axis=1)
			self.df['куда'] = self.df.loc[:, ['Получатель.Адрес.Город', 'Получатель.Адрес']].apply(Readdf.ower_city, axis=1)