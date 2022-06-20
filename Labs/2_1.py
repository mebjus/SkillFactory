import pandas as pd
from pandas import DataFrame

melb_data = pd.read_csv('data/melb_data_fe.csv', sep=',')
melb_df = melb_data.copy()

melb_df['quarter'] = pd.to_datetime(melb_df['Date']).dt.quarter
# print(melb_df['quarter'].value_counts())


exception_w = ('Date', 'Rooms', 'Bedroom', 'Bathroom', 'Car')
for col in melb_df.columns:
    if (melb_df[col].nunique() < 150) and (col not in exception_w):
        melb_df[col] = melb_df[col].astype('category')

# print(melb_df.sort_values(by=['AreaRatio'], ascending=False, ignore_index=True).loc[1558, ['AreaRatio','BuildingArea']])

f1 = melb_df['Rooms'] > 2
f2 = melb_df['Type'] == 'townhouse'

# print(melb_df[f1 & f2].sort_values(by=['Rooms','MeanRoomsSquare'], ascending=[True, False], ignore_index=True).loc[18,['Price']])
# print(melb_df.groupby('Rooms')['Price'].mean().sort_values(ascending=False))
# print(melb_df.groupby('Regionname')['Lattitude'].std().sort_values(ascending=False))

date1 = pd.to_datetime('2017-05-01')
date2 = pd.to_datetime('2017-09-01')
t = pd.to_datetime(melb_df['Date'])
mask = (date1 <= t) & (t <= date2)

# print(melb_df[mask].groupby('SellerG')['Price'].sum().sort_values(ascending=True))

# print(melb_df.groupby(['Rooms', 'Type'])['Price'].mean())
# print(melb_df.groupby(['Rooms', 'Type'])['Price'].mean().unstack())

pivot: DataFrame = melb_df.pivot_table(values='BuildingArea', index='Type', columns='Rooms', aggfunc='median')
pivot2 = melb_df.pivot_table(values='Price', index='SellerG', columns='Type', aggfunc='mean')
mask2 = pivot2['unit'].max()
print(pivot2[pivot2['unit'] == mask2])
