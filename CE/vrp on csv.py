import pandas as pd
import datetime

df = pd.read_csv('data/DispatchesCount.csv')
df.columns = ['город', 'дата', 'шт', 'цена паблик', 'цена', 'кг', '1', '2', '3', '4']
df.drop('1', axis=1, inplace=True)
df.drop('2', axis=1, inplace=True)
df.drop('3', axis=1, inplace=True)
df.drop('4', axis=1, inplace=True)
df = df.fillna(0)


# for i in range(0,10):
#     df.drop(i, axis=0)

def todate(arg):
    arg = pd.to_datetime(arg)
    return arg.strftime('%Y-%m')


# df['дата'] = df['дата'].apply(todate)
# df['дата'] = pd.to_datetime(df['дата'], format='%m%Y')

# df = df.groupby('дата', as_index=False).aggregate(sum)

print(df['дата'])

dict_fo = {'СЗФО': ['Великий Новгород', 'Мурманск', 'Петрозаводск', 'Сыктывкар', 'Санкт-Петербург', 'Архангельск',
                    'Калининград'],
           'УФО': ['Курган', 'Нижневартовск', 'Новый Уренгой', 'Стерлитамак', 'Магнитогорск', 'Оренбург', 'Сургут',
                   'Екатеринбург', 'Пермь', 'Тюмень', 'Уфа', 'Челябинск'],
           'ПФО': ['Ижевск', 'Пенза', 'Ульяновск', 'Чебоксары', 'Киров', 'Нижний Новгород', 'Казань', 'Самара',
                   'Саратов', 'Тольятти'],
           'ЮФО': ['Новороссийск', 'Симферополь', 'Пятигорск', 'Ростов-на-Дону', 'Волгоград', 'Воронеж', 'Краснодар',
                   'Ставрополь',
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

# df['ФО'] = df['город'].apply(ret)


# dfg = df.pivot_table(index='ФО', columns='дата', values=['шт','цена','кг'], aggfunc='sum')
# ['шт','цена','кг']

