import pandas as pd

byke = pd.read_csv('data/citibike-tripdata.csv', sep=',')
byke_df = byke.copy()
# g = byke_df['start station id'].value_counts() #759
# g1 = byke_df['end station id'].value_counts() #765

# byke_df['Time'] = pd.to_datetime(byke_df['starttime'])
# age = (byke_df['Time'].dt.year - byke_df['birth year'])
# print(age.min())


# g = byke_df['end station name'].value_counts()
# print(g)

byke_df.drop(['start station id', 'end station id'], axis=1, inplace=True)
byke_df['age'] = pd.to_datetime(byke_df['starttime'])
byke_df['age'] = byke_df['age'].dt.year - byke_df['birth year']
byke_df.drop(['birth year'], axis=1, inplace=True)
g = byke_df['age'] > 60

byke_df['start'] = pd.to_datetime(byke_df['starttime'])
byke_df['stop'] = pd.to_datetime(byke_df['stoptime'])
byke_df['trip duration'] = byke_df['stop'] - byke_df['start']


# print(byke_df['trip duration'].dt.seconds.mean())

def is_weekend(df):
    if df == 5 or df == 6:
        return 1
    else:
        return 0


byke_df['weekend'] = byke_df['start'].dt.weekday.apply(is_weekend)
print(byke_df['weekend'].value_counts())

def is_time_of_day(df):
    c = None
    if 0 <= df <= 6:
        c = 'night'
    elif 6 < df <= 12:
        c = 'morning'
    elif 12 < df <= 18:
        c = 'day'
    elif 18 < df <= 23:
        c = 'evening'
    return c

byke_df['time_of_day'] = byke_df['start'].dt.hour.apply(is_time_of_day)
print(byke_df['time_of_day'].value_counts())

