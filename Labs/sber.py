import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
import seaborn as sns

sber_data = pd.read_csv('data/sber_data.csv')
# print(sber_data.shape)
# fig = plt.figure(figsize=(10, 5))

# sns.scatterplot(data=sber_data, x='price_doc', y='kremlin_km', size=1)
# plt.show()

# cols_null_percent = sber_data.isnull().mean() * 100
# cols_with_null = cols_null_percent[cols_null_percent>0].sort_values(ascending=False)

# drop_data = sber_data.copy()
# thresh = drop_data.shape[0]*0.7
# drop_data = drop_data.dropna(how='any', thresh=thresh, axis=1)
# drop_data = drop_data.dropna(how='any', axis=0)
# drop_data.isnull().mean()


# cols_null_percent = sber_data.isnull().mean() * 100
# cols_with_null = cols_null_percent[cols_null_percent>0].sort_values(ascending=False)
# cols_with_null.plot(
#     kind='bar',
#     figsize=(10, 4),
#     title='Распределение пропусков в данных'
# )
# # combine_data = sber_data.copy()
#
# #отбрасываем столбцы с числом пропусков более 30% (100-70)
# n = combine_data.shape[0] #число строк в таблице
# thresh = n*0.7
# combine_data = combine_data.dropna(how='any', thresh=thresh, axis=1)
#
# #отбрасываем строки с числом пропусков более 2 в строке
# m = combine_data.shape[1] #число признаков после удаления столбцов
# combine_data = combine_data.dropna(how='any', thresh=m-2, axis=0)

#создаём словарь 'имя_столбца': число (признак), на который надо заменить пропуски
# values = {
#     'life_sq': combine_data['full_sq'],
#     'metro_min_walk': combine_data['metro_min_walk'].median(),
#     'metro_km_walk': combine_data['metro_km_walk'].median(),
#     'railroad_station_walk_km': combine_data['railroad_station_walk_km'].median(),
#     'railroad_station_walk_min': combine_data['railroad_station_walk_min'].median(),
#     'preschool_quota': combine_data['preschool_quota'].mode()[0],
#     'school_quota': combine_data['school_quota'].mode()[0],
#     'floor': combine_data['floor'].mode()[0]
# }
#заполняем оставшиеся записи константами в соответствии со словарем values
# combine_data = combine_data.fillna(values)
# plt.show()

# fig, axes = plt.subplots(nrows=1, ncols=2, figsize=(15, 4))
# histplot = sns.histplot(data=sber_data, x='full_sq', ax=axes[0]);
# histplot.set_title('Full Square Distribution');
# boxplot = sns.boxplot(data=sber_data, x='full_sq', ax=axes[1]);
# boxplot.set_title('Full Square Boxplot')
#
# plt.show()

def outliers_iqr(data, feature, log_scale=False, left=1.5, right=1.5):
    if log_scale:
        x = np.log(data[feature]+1)
    else:
        x = data[feature]
    quartile_1, quartile_3 = x.quantile(0.25), x.quantile(0.75),
    iqr = quartile_3 - quartile_1
    lower_bound = quartile_1 - (iqr * left)
    upper_bound = quartile_3 + (iqr * right)
    outliers = data[(x<lower_bound) | (x > upper_bound)]
    cleaned = data[(x>lower_bound) & (x < upper_bound)]
    return outliers, cleaned
#
# outliers, cleaned = outliers_iqr(sber_data, 'full_sq', 1, 6)
# print(f'Число выбросов по методу Тьюки: {outliers.shape[0]}')
# print(f'Результирующее число записей: {cleaned.shape[0]}')
# fig, axes = plt.subplots(nrows=1, ncols=2, figsize=(15, 4))
# histplot = sns.histplot(data=cleaned, x='full_sq', ax=axes[0]);
# histplot.set_title('Cleaned Full Square Distribution');
# boxplot = sns.boxplot(data=cleaned, x='full_sq', ax=axes[1]);
# boxplot.set_title('Cleaned Full Square Boxplot')
# plt.show()

# ig, axes = plt.subplots(1, 2, figsize=(15, 4))

# histplot = sns.histplot(sber_data['mkad_km'], bins=30, ax=axes[0])
# histplot.set_title('MKAD Km Distribution');
#
# #гистограмма в логарифмическом масштабе
log_mkad_km= np.log(sber_data['mkad_km'] + 1)
# histplot = sns.histplot(log_mkad_km , bins=30, ax=axes[1])
# histplot.set_title('Log MKAD Km Distribution')
# plt.show()
# print(log_mkad_km.skew())

def outliers_z_score(data, feature, log_scale=False, left=3, right=3):
    if log_scale:
        x = np.log(data[feature]+1)
    else:
        x = data[feature]
    mu = x.mean()
    sigma = x.std()
    lower_bound = mu - left * sigma
    upper_bound = mu + right * sigma
    outliers = data[(x < lower_bound) | (x > upper_bound)]
    cleaned = data[(x > lower_bound) & (x < upper_bound)]
    return outliers, cleaned
#
# outliers, cleaned = outliers_iqr(sber_data, 'price_doc', log_scale=True, left=3, right=3)
# print(f'Число выбросов по методу z-отклонения: {outliers.shape[0]}')
# print(f'Результирующее число записей: {cleaned.shape[0]}')
#
# fig, ax = plt.subplots(1, 1, figsize=(8, 4))
# log_price = np.log(sber_data['price_doc'] + 1)
# histplot = sns.histplot(log_price, bins=30, ax=ax)
# histplot.axvline(log_price.mean(), color='r', lw=2)
# histplot.axvline(log_price.mean()+ 3 * log_price.std(), color='k', ls='--', lw=2)
# histplot.axvline(log_price.mean()- 3 * log_price.std(), color='k', ls='--', lw=2)
# # histplot.set_title('Log MKAD Km Distribution')
# plt.show()

# dupl_columns = list(sber_data.columns)
# dupl_columns.remove('id')
#
# mask = sber_data.duplicated(subset=dupl_columns)
# sber_duplicates = sber_data[mask]
# sber_dedupped = sber_data.drop_duplicates(subset=dupl_columns)
# print(f'Результирующее число записей: {sber_dedupped.shape[0]}')
#
# low_information_cols = []
#
# #цикл по всем столбцам
# for col in sber_data.columns:
#     top_freq = sber_data[col].value_counts(normalize=True).max()
#     nunique_ratio = sber_data[col].nunique() / sber_data[col].count()
#     if top_freq > 0.95:
#         low_information_cols.append(col)
#         print(f'{col}: {round(top_freq*100, 2)}% одинаковых значений')
#     if nunique_ratio > 0.95:
#         low_information_cols.append(col)
#         print(f'{col}: {round(nunique_ratio*100, 2)}% уникальных значений')
#
# information_sber_data = sber_data.drop(low_information_cols, axis=1)
# print(f'Результирующее число признаков: {information_sber_data.shape[1]}')
#

data = pd.read_excel('data/ratings+movies.xlsx')
print(data)
data1= pd.read_excel('data/ratings+movies.xlsx', sheet_name='movies')
data = data.merge(data1, how='left')
print(data)
