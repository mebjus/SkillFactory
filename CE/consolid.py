import pandas as pd
import os

df = pd.DataFrame
dirname = 'data/kis/'
dirfiles = os.listdir(dirname)
fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
pd.options.display.float_format = '{:,.2F}'.format

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


# def todate(arg):
#     arg = pd.to_datetime(arg)
#     return arg.strftime('%Y-%m-%d')
#
#
def todate1(arg):
    arg = pd.to_datetime(arg)
    return arg.strftime('%Y-%m')
#
#
# df['Дата Cоздания'] = df['Дата Cоздания'].apply(todate)

df['Дата Cоздания'] = df['Дата Cоздания'].dt.strftime('%Y-%m-%d')

df.rename(columns={'Дата Cоздания': 'дата',
                   'Номер отправления': 'шт', 'Общая стоимость со скидкой': 'деньги'},
          inplace=True)

# отбрасываем все нулевки, консолидированные сборы, дешевые доборы
# df = df[df['деньги'] > 10]

## преобразуем  дата фрейм под Москву

df.loc[:,'шт'] = df.loc[:, 'шт'].apply(lambda x: x[0:12])   # отрезаем отправления -
print(df['шт'])
mask1 = df['Отправитель.Адрес.Город'] == 'Москва'
mask2 = df['Получатель.Адрес.Город'] == 'Москва'
df_dep = df[mask1]
df_arr = df[mask2]

df_inner = df[(mask1 & mask2)]
df_or = df[(mask1 | mask2)]


df_cons_dep = df_dep.groupby(['шт', 'дата'])['Клиент'].count().reset_index()
df_cons_arr = df_arr.groupby(['шт', 'дата'])['Клиент'].count().reset_index()


df_cons_dep.rename(columns={'шт': 'номер', 'Клиент': 'количество'}, inplace=True)
df_cons_arr.rename(columns={'шт': 'номер', 'Клиент': 'количество'}, inplace=True)

df_cons_dep = df_cons_dep[df_cons_dep['количество'] > 1]
df_cons_arr = df_cons_arr[df_cons_arr['количество'] > 1]

count_all = df_dep.shape[0] + df_arr.shape[0]
count_stop = count_all - df_cons_dep['количество'].sum() + df_cons_dep['номер'].count()

print('Кол отправлений из Москвы или в Москву: {:,.0F}'.format(count_all))
print('Кол отправлений в Москву из Москвы (вкл в общ): ', df[(mask1 & mask2)].shape[0])
print('Кол стопов: {:,.0F}'.format(count_stop))

count_rd = df_or['дата'].apply(todate1).reset_index()
count_rd = count_rd.groupby('дата').sum()
count_rd = count_rd.reset_index()
count_rd['р.д.'] = count_rd['дата'].apply(lambda x: mounth[str(x)])


print('Сумма деньги: {:,.0F}'.format((df_inner['деньги'].sum() + df_or['деньги'].sum())/2))
print('Расчетный вес: {:,.0F}'.format((df_inner['Расчетный вес'].sum() + df_or['Расчетный вес'].sum())/2))


df = pd.Series({'Количество отправлений': count_all, 'Количество стопов': count_stop,
                'Количество рд': count_rd['р.д.'][0]})

print('Количество рд', count_rd['р.д.'][0])


##################### сохраняем в файл

# writer = pd.ExcelWriter('consolid.xlsx', engine='xlsxwriter')
#
# df.to_excel(writer, sheet_name='итоги', header=False, index=True)
# df_cons_dep.to_excel(writer, sheet_name='отправлено из Мск', index=False)  # index=False header=False
# df_cons_arr.to_excel(writer, sheet_name='доставлено в Мск', index=False)
#
# df_or.to_excel(writer, sheet_name='отправитель_получатель', index=False)  # index=False header=False
# df_inner.to_excel(writer, sheet_name='из Мск в Мск', index=False)
#
# workbook = writer.book
# writer.save()
