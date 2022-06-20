import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns


sns.set_style("darkgrid")

churn_data = pd.read_csv('data/churn.csv')
df = churn_data.copy()
df = df.drop(['RowNumber'], axis=1)

# print(df.info())

# fig = plt.figure(figsize=(6, 5))
# ex = df['Exited'].value_counts()
# plt.pie(ex, labels = ['Лояльные','Ушедшие'], autopct='%.0f%%', explode=[0.1, 0])
# plt.suptitle('Диаграмма, показывающая соотношение ушедших и лояльных клиентов')
# plt.show()
# Соотношение клиентов 80/20 в пользу лояльных, для понимания хорошо это или плохо необходимо сравнение
# либо по рынку с конкурентами, либо в исторической хронологии. Для неспециалиста в области - звучит
# прекрасно

# 9.2. Постройте график, показывающий распределение баланса пользователей,
# у которых на счету больше 2 500 долларов. Опишите распределение и сделайте выводы.

df_bal = df[df['Balance'] > 2500].reset_index(drop=True)
# print(df_bal)

sns.displot(x='Balance',data=df_bal);
plt.show()
