import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_csv('data/diabetes_data.csv')
# print(df.shape)
# fig, axes = plt.subplots(nrows=1, ncols=5, figsize=(15, 3))
# sns.histplot(data=df, x='BloodPressure', ax=axes[0])
# sns.histplot(data=df, x='BMI', ax=axes[1])
# sns.histplot(data=df, x='Pregnancies', ax=axes[2])
# sns.histplot(data=df, x='SkinThickness', ax=axes[3])
# sns.histplot(data=df, x='DiabetesPedigreeFunction', ax=axes[4])
# plt.show()

dupl = list(df.columns)
# mask = df.duplicated(subset=dupl)
# df_dupl = df[mask]
# print(df_dupl)

df = df.drop_duplicates(subset=dupl)

low_information_cols = []
for col in df.columns:
    top_freq = df[col].value_counts(normalize=True).max()
    nunique_ratio = df[col].nunique() / df[col].count()
    if top_freq > 0.99:
        low_information_cols.append(col)
        # print(f'{col}: {round(top_freq*100, 2)}% одинаковых значений')
    if nunique_ratio > 0.99:
        low_information_cols.append(col)
        # print(f'{col}: {round(nunique_ratio*100, 2)}% уникальных значений')

df = df.drop(low_information_cols, axis=1)
# print(f'Результирующее число признаков: {inf.shape[1]}')

nan_list=['BloodPressure', 'Glucose', 'SkinThickness', 'Insulin', 'BMI']

def nan_change(cell):
    if cell == 0:
        return np.nan
    else:
        return cell


for i in nan_list:
    df[i] = df[i].apply(nan_change)

# print(df['Insulin'].isnull().mean().round(2))

low_information_cols=[]

thresh = df.shape[0]*0.7
df = df.dropna(thresh=thresh, axis=1)
# print(df.shape[1])

m = df.shape[1] #число признаков после удаления столбцов
df = df.dropna(how='any', thresh=m-2, axis=0)

# df = df.drop(low_information_cols, axis=1)
# print(f'Результирующее число признаков: {df.shape[1]}')

values = {
    'Pregnancies': df['Pregnancies'].median(),
    'Glucose': df['Glucose'].median(),
    'BloodPressure': df['BloodPressure'].median(),
    'SkinThickness': df['SkinThickness'].median(),
    'DiabetesPedigreeFunction': df['DiabetesPedigreeFunction'].median(),
    'Age': df['Age'].median(),
    'Outcome': df['Outcome'].median(),
}
# заполняем оставшиеся записи константами в соответствии со словарем values
df = df.fillna(values)
# print(df['SkinThickness'].mean())

def outliers_iqr(data, feature, log_scale=False, left=1.5, right=1.5):
    if log_scale:
        x = np.log(data[feature])
    else:
        x = data[feature]
    quartile_1, quartile_3 = x.quantile(0.25), x.quantile(0.75),
    iqr = quartile_3 - quartile_1
    lower_bound = quartile_1 - (iqr * left)
    upper_bound = quartile_3 + (iqr * right)
    outliers = data[(x<lower_bound) | (x > upper_bound)]
    cleaned = data[(x>lower_bound) & (x < upper_bound)]
    return outliers

# outliers, cleaned = outliers_iqr(df, 'DiabetesPedigreeFunction')
# print(f'Число выбросов по методу Тьюки: {outliers.shape[0]}')
# print(f'Результирующее число записей: {cleaned.shape[0]}')

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
    return outliers

# outliers, cleaned = outliers_z_score(df, 'DiabetesPedigreeFunction')
# print(f'Число выбросов по методу z-отклонения: {outliers.shape[0]}')


o1 = outliers_iqr(df, 'DiabetesPedigreeFunction')
o2 = outliers_iqr(df, 'DiabetesPedigreeFunction', log_scale=True)
print(o1.shape[0])
print(o2.shape[0])
