import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel('data/DispatchesCount_online.xlsx', header=1)
# df = pd.read_csv('data/DispatchesCount.csv', header=1)
df_m = pd.read_excel('data/day_of_month.xlsx')

df = df.drop('Unnamed: 7', axis=1)
df = df.fillna(0)

for i in range(1, df.shape[0], 4):
    df.rename(index={i: df.iloc[i][0] + ' шт',
                     i + 1: df.iloc[i][0] + ' без скидки',
                     i + 2: df.iloc[i][0] + ' деньги',
                     i + 3: df.iloc[i][0] + ' кг',
                     }, inplace=True)

df = df.T

def todate(arg):
    y = pd.to_datetime(arg)
    y = y.strftime('%Y-%m')
    return y


df[0] = df[0].apply(todate)

low = df[df['Итого шт'] == 'Итого']
df = df.drop(low.index, axis=0)

df.rename(columns={0: 'Дата'}, inplace=True)

df = df.groupby(['Дата'], as_index=False).sum().round(0)
df = df.merge(df_m, how='left')

df['Итого, деньги на рд'] = (df['Итого деньги'] / df['count']).round()
df['Итого, шт на рд'] = (df['Итого шт'] / df['count']).round()
df['Итого, кг на рд'] = (df['Итого кг'] / df['count']).round()

fig = plt.figure(figsize=(15, 5))

plt.subplot(131)
sns.lineplot(x = 'Дата', y='Итого, кг на рд', data=df, lw=5)
plt.xticks(rotation=45)

plt.subplot(132)
sns.lineplot(x = 'Дата', y='Итого, деньги на рд', data=df, lw=5)
plt.xticks(rotation=45)

plt.subplot(133)
sns.lineplot(x = 'Дата', y='Итого, шт на рд', data=df, lw=5)
plt.xticks(rotation=45)
plt.show()

can = ['Дата', 'Итого шт', 'Итого деньги', 'Итого кг', 'count', 'Итого, деньги на рд', 'Итого, кг на рд',
       'Итого, шт на рд']
for i in df.columns:
    if i not in can:
        df.drop(i, axis=1, inplace=True)

df.to_excel('./summary.xlsx', sheet_name='итоги', index=False)
