import pandas as pd

data = pd.read_excel('C:\\temp\\SkillFactory\\CE\\18_11.xls')
mask = data['Общая стоимость со скидкой'] > 25000
s = data.groupby('Клиент')['Общая стоимость со скидкой'].sum().sort_values(ascending=False)
# print(s)
# print(data[mask])
s.to_csv('out.csv')
# data[mask].to_csv('out.csv')
# надо проверить как работает
# или нет