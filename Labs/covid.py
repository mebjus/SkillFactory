import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly
import plotly.express as px
import seaborn as sns



covid_data = pd.read_csv('data/covid/covid_data.csv')
vaccinations_data = pd.read_csv('data/covid/country_vaccinations.csv')
vaccinations_data = vaccinations_data[
    ['country', 'date', 'total_vaccinations',
     'people_vaccinated', 'people_vaccinated_per_hundred',
     'people_fully_vaccinated', 'people_fully_vaccinated_per_hundred',
     'daily_vaccinations', 'vaccines']
]


covid_data = covid_data.groupby(['date', 'country'], as_index=False)[['confirmed', 'deaths', 'recovered']].sum()

covid_data['date'] = pd.to_datetime(covid_data['date'])
covid_data['active'] = covid_data['confirmed'] - covid_data['deaths'] - covid_data['recovered']
covid_data = covid_data.sort_values(by=['country', 'date'])

covid_data['daily_confirmed'] = covid_data.groupby('country')['confirmed'].diff()
covid_data['daily_deaths'] = covid_data.groupby('country')['deaths'].diff()
covid_data['daily_recovered'] = covid_data.groupby('country')['recovered'].diff()

vaccinations_data['date'] = pd.to_datetime(vaccinations_data['date'])

# print(covid_data['date'].min(), covid_data['date'].max())

covid_df = covid_data.merge(vaccinations_data, on=['date', 'country'], how='left')
# print(covid_df.shape)

covid_df['death_rate'] = (covid_df['deaths'] / covid_df['confirmed']) * 100
covid_df['recover_rate'] = (covid_df['recovered'] / covid_df['confirmed']) * 100

c = covid_df[covid_df['country'] == 'United States'].groupby('country')['death_rate'].max()
# print(round(c, 2))
russia = covid_df[covid_df['country'] == 'Russia'].groupby('country')['recover_rate'].mean()
# print(round(russia, 2))
grouped_cases = covid_df.groupby('date')['daily_confirmed'].sum()
# grouped_cases.plot(kind='line', figsize=(12, 4), title='Ежедневная заболеваемость по всем странам', grid=True, lw=3)
# grouped_cases.plot(
#     kind='hist',
#     figsize=(10, 6),
#     title='Распределение ежедневной заболеваемости',
#     grid = True,
#     color = 'black',
#     bins=10
# )

# grouped_country = covid_df.groupby(['country'])['confirmed'].last()
# grouped_country = grouped_country.nlargest(10)
# grouped_country.plot(
#     kind='bar',
#     grid=True,
#     figsize=(12, 4),
#     colormap='plasma'
# )

# grouped_country = covid_df.groupby(['country'])[['confirmed', 'deaths']].last()
# grouped_country = grouped_country.nlargest(10, columns=['confirmed'])
# grouped_country.plot(
#     kind='bar',
#     grid=True,
#     figsize=(12, 4),
# );

# covid_df.groupby(['country'])['total_vaccinations'].last().nsmallest(5).plot(kind='bar')
# plt.show()

# line_data = covid_df.groupby('date', as_index=False).sum()
# fig = px.line(
#     data_frame=line_data, #DataFrame
#     x='date', #ось абсцисс
#     y='daily_vaccinations', #ось ординат
#     height=500, #высота
#     width=1000, #ширина
#     title='Confirmed, Recovered, Deaths, Active cases over Time' #заголовок
# )
# fig.show()

choropleth_data = vaccinations_data.sort_values(by='date')
choropleth_data['date'] = choropleth_data['date'].astype('string')
# print(choropleth_data.info())
#строим график
# fig = px.choropleth(
#     data_frame=choropleth_data, #DataFrame
#     locations="country", #столбец с локациями
#     locationmode = "country names", #режим сопоставления локаций с базой Plotly
#     color="total_vaccinations", #от чего зависит цвет
#     animation_frame="date", #анимационный бегунок
#     range_color=[0, 30e6], #диапазон цвета
#     title='Global Spread of COVID-19', #заголовок
#     width=800, #ширина
#     height=500, #высота
#     color_continuous_scale='Reds' #палитра цветов
# )
#
# # отображаем график
# fig.show()

# fig, ax = plt.subplots()
# offset=0.4
# data = np.array([[5, 10, 7], [8, 15, 5], [11, 9, 7]])
# cmap = plt.get_cmap('tab20b')
# b_colors = cmap(np.array([0, 8, 12]))
# sm_colors = cmap(np.array([1, 2, 3, 9, 10, 11, 13, 14, 15]))
# ax.pie(data.sum(axis=1), radius=1, colors=b_colors,
# wedgeprops=dict(width=offset, edgecolor='w'))
# ax.pie(data.flatten(), radius=1-offset, colors=sm_colors,
# wedgeprops=dict(width=offset, edgecolor='w'))


sns.set_style("darkgrid")
np.random.seed(123)

sample = [i for i in range(10)]
y = np.random.randint(10, size=len(sample))
z = np.random.randint(4, size=len(sample))
data = [sample, y, z]


df = pd.DataFrame(data).transpose()
df.columns = ['sample', 'y_val', 'z_val']
sns.lineplot(data=df)

plt.show()
