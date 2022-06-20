import pandas as pd
from pandas_profiling import ProfileReport
import sweetviz as sv
import statistics as st
import matplotlib.pyplot as plt
from scipy import stats
# from sklearn.metrics import matthews_corrcoef
import seaborn as sns


data = pd.read_csv('wine.csv')
# print(data['price'].max())
# print(data[taster_name].nunique())
# print(data.info())


# sweetviz
# report = sv.analyze(data)
# report.show_html()

# profile = ProfileReport(data, title="Wine Pandas Profiling Report")
# print(profile)

# ищем дубликаты
dupl = list(data.columns)
data_dupl = data[data.duplicated(subset=dupl)]
# print(data_dupl.shape[0])
data_dupl.index.nunique()
data = data.drop_duplicates(subset=dupl)

# ищем пропуски
# print(data.isnull().sum())

# #отбрасываем столбцы с числом пропусков более 30% (100-70)
n = data.shape[0] #число строк в таблице
thresh = n*0.7
data = data.dropna(how='any', thresh=thresh, axis=1)
# print(data.isnull().sum())

values = {
    'country': 'unknown',
    'designation': 'unknown',
    'price': -1,
    'province': 'unknown',
    'region_1': 'unknown',
    'taster_name': 'unknown',
    'taster_twitter_handle': 'unknown',
    'variety': 'unknown'}
data = data.fillna(values)


# К таким инструментам можно отнести следующие библиотеки Python,
# которые могут выполнять EDA всего одной строкой кода:
#
# d-tale;
# pandas-profiling;
# sweetviz.


df = pd.read_csv('wine_cleared.csv')

# print(df.corr(method = 'kendall'))

df = pd.read_csv('model.csv')
print(df.corr())
# sns.heatmap(df.corr(), annot = True)
# plt.show()

report = sv.analyze(df)
report.show_html()