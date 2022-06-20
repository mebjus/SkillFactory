import pandas as pd
import category_encoders as ce
from scipy.stats import shapiro
from scipy.stats import normaltest
from scipy.stats import ttest_ind
from scipy.stats import spearmanr
from scipy.stats import f_oneway # ANOVA


petersburg = [0.0974, 0.1352, 0.0817, 0.1016, 0.0968, 0.1064, 0.105]
magadan = [0.1033, 0.0915, 0.0781, 0.0685, 0.0677, 0.0697, 0.0764,
           0.0689]
           
df = pd.concat([pd.Series(petersburg), pd.Series(magadan)], axis=1)
df.rename(columns={0: 'petersburg', 1: 'magadan'}, inplace=True)

df['petersburg'][7] = round(df['petersburg'].median(), 4)

print('отклонение в среднем:', round(df['petersburg'].mean()-df['magadan'].mean(), 3))


# проверяем на нормальность распределения
H0 = 'Данные распределены нормально'
Ha = 'Данные не распределены нормально (мы отвергаем H0)'

alpha = 0.05
_, p = shapiro(df)
print(round(p,4))

# Интерпретация

if p > alpha:
	print(H0)
else:
	print(Ha)

print(df.corr())

print('Корреляция - очень слабая или даже отсутствует, метод корреляции Пирсона, тк выяснили что распределение нормальное, без выбросов')

H0 = 'Нет значимой разницы между средним размером раковины мидий в двух разных местах.'
Ha = 'Есть значимая разница между средним размером раковины мидий в двух разных местах.'

def t_test(df):
    print('\n' + "*** Результаты независимого T-теста ***")
    test_results = ttest_ind(df['petersburg'], df['magadan'], equal_var=True)

    p = round(test_results[1],3)
    print(p)

    if p>alpha:
        print(f"{p} > {alpha}. \n Мы не можем отвергнуть нулевую гипотезу. \n {H0}")
    else:
        print(f"{p} <= {alpha}. \n Мы отвергаем нулевую гипотезу. \n {Ha}")

t_test(df)