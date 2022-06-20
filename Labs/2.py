import pandas as pd
import numpy as np

# def create_companyDF(income, expenses, years):
#     """
#     Создайте функцию create_companyDF(income, expenses, years), которая  возвращает DataFrame, 
#     составленный из входных данных со столбцами “Income” и “Expenses” и индексами, соответствующим годам рассматриваемого периода.
#     """
#     df = pd.DataFrame({'Income': income, 'Expenses': expenses}, index = years)
#     return df

# def get_profit(df, year):

#     """
#     А также напишите функцию get_profit(df, year), которая возвращает разницу между доходом и расходом, записанных в таблице df, за год year.
#     Учтите, что если информация за запрашиваемый год не указана в вашей таблице вам необходимо вернуть None. 
#     """
#     profit = df[year]
#     return profit

# expenses = [156, 130, 270]
# income = [478, 512, 196]
# years = [2018, 2019, 2020]

# scienceyou = create_companyDF(income, expenses, years)
# print(scienceyou)
# print(scienceyou.loc[2019]['Income'])
# # print(get_profit(scienceyou, 2019))


# ufo_data = pd.read_csv('data/ufo.csv', sep=',')
#
#
# ufo_data['Data'] = pd.to_datetime(ufo_data['Time'])
# data_vision = ufo_data['Data'].dt.date
# ufo_NV = ufo_data[ufo_data['State'] == 'NV']
# r = ufo_NV['Data'].diff()
# r = r.dt.days
# print(round(r.mean()))

melb_data = pd.read_csv('data/melb_data_ps.csv', sep=',')
melb_df = melb_data.copy()
melb_df.drop(['index', 'Coordinates'], axis=1, inplace=True)
total_rooms = melb_df['Rooms'] + melb_df['Bedroom'] + melb_df['Bathroom']
melb_df['MeanRoomSquare'] = melb_df['BuildingArea'] / total_rooms
diff_area = melb_df['BuildingArea'] - melb_df['Landsize']
sum_area = melb_df['BuildingArea'] + melb_df['Landsize']
melb_df['AreaRatio'] = diff_area/sum_area

melb_df['Date'] = pd.to_datetime(melb_df['Date'])
melb_df['AgeBuilding'] = melb_df['Date'].dt.year - melb_df['YearBuilt']


melb_df['d'] = melb_df['Date'].dt.dayofweek
d = melb_df['d'].value_counts()


def weekday(WeekdaySale):
    if WeekdaySale == 5 or WeekdaySale == 6:
        return 1
    else:
        return 0

# def get_street_type(address):
#     exclude_list = ['N', 'S', 'W', 'E']
#     address_list = address.split(' ')
#     street_type = address_list[-1]
#     if street_type in exclude_list:
#         street_type = address_list[-2]
#     return street_type
#
# street_types = melb_df['Address'].apply(get_street_type)
# popular_stypes =street_types.value_counts().nlargest(10).index

melb_df['Weekend'] = melb_df['d'].apply(weekday)
print(melb_df['Weekend'].value_counts())

# k = melb_df[melb_df['Weekend'] == 1]
# print(k['Price'].mean())

# seller = melb_df['SellerG']
# popular_seller = melb_df['SellerG'].value_counts().nlargest(49).index
# melb_df['Pop_seller'] = seller.apply(lambda x: x if x in popular_seller else 'other')
# k = melb_df[melb_df['Pop_seller'] == 'other']['Price'].min()
# l = melb_df[melb_df['Pop_seller'] == 'Nelson']['Price'].min()
# print(round(l/k, 1))


# memory usage: 2.3+ MB
# suburb = melb_df['Suburb']
# popular_suburb = melb_df['Suburb'].value_counts().nlargest(119).index
# melb_df['Suburb'] = suburb.apply(lambda x: x if x in popular_suburb else 'other')
# melb_df['Suburb'] = melb_df['Suburb'].astype('category')
# print(melb_df.info())
