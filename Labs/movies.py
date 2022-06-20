import pandas as pd

movies = pd.read_csv('data/movies.csv', sep=',')
rating = pd.read_csv('data/ratings1.csv', sep=',')
rating2 = pd.read_csv('data/ratings2.csv', sep=',')
dates = pd.read_csv('data/dates.csv', sep=',')


# print(movies.nunique())
# print(rating.nunique())
# print(pd.to_datetime(dates1['date']).dt.year.mode())

ratings = pd.concat([rating, rating2], ignore_index=True)
ratings = ratings.drop_duplicates(ignore_index=True)

ratings_dates = pd.concat([ratings, dates], axis=1)
joined_false = ratings_dates.join(movies, rsuffix='_right', how='inner')
joined = ratings_dates.join(movies.set_index('movieId'), on='movieId', how='left')
merged = ratings_dates.merge(movies, on='movieId', how='left')
# print(joined)

data_1 = pd.DataFrame({'Value': [100, 45, 80],
                       'Group': [1, 4, 5]},
                      index=['I0', 'I1', 'I2']
                      )
data_2 = pd.DataFrame({'Company': ['Google', 'Amazon', 'Facebook'],
                       'Add': ['S0', 'S1', 'S7']},
                      index=['I0', 'I1', 'I3'])

# print(data_1.join(data_2, how='inner'))
# print(data_1.merge(data_2, how='inner', right_index=True, left_index=True))

# items_df = pd.DataFrame({
#     'item_id': [417283, 849734, 132223, 573943, 19475, 3294095, 382043, 302948, 100132, 312394],
#     'vendor': ['Samsung', 'LG', 'Apple', 'Apple', 'LG', 'Apple', 'Samsung', 'Samsung', 'LG', 'ZTE'],
#     'stock_count': [54, 33, 122, 18, 102, 43, 77, 143, 60, 19]
# })
#
# purchase_df = pd.DataFrame({
#     'purchase_id': [101, 101, 101, 112, 121, 145, 145, 145, 145, 221],
#     'item_id': [417283, 849734, 132223, 573943, 19475, 3294095, 382043, 302948, 103845, 100132],
#     'price': [13900, 5330, 38200, 49990, 9890, 33000, 67500, 34500, 89900, 11400]
# })
# print(items_df)
# print(purchase_df)
#
# merged = items_df.merge(purchase_df, how='inner')
# merged['total'] = merged['stock_count'] * merged['price']
# income = merged['total'].sum()
# print(income)

ratings = pd.read_csv('data/ratings_movies.csv', sep=',')
import re
def get_year_release(arg):
    #находим все слова по шаблону "(DDDD)"
    candidates = re.findall(r'\(\d{4}\)', arg)
    # проверяем число вхождений
    if len(candidates) > 0:
        #если число вхождений больше 0,
	#очищаем строку от знаков "(" и ")"
        year = candidates[0].replace('(', '')
        year = year.replace(')', '')
        return int(year)
    else:
        #если год не указан, возвращаем None
        return None


ratings['year'] = ratings['title'].apply(get_year_release)
# print(ratings['year'].info)

ratings['year'] = ratings['title'].apply(get_year_release)
# #  userId   movieId    rating  date    title   genres   year
# mask = ratings['year'] == 2018
# c = (ratings[mask].groupby(by='genres')['rating'].agg(['mean', 'count']))
# print(c.mean)

# mask = ratings['year'] == 2018
# grouped = ratings[mask].groupby('genres')['rating'].agg(['mean', 'count'])
# print(grouped[grouped['count']>10].sort_values(by='mean',ascending=False))

# ratings['year'] = pd.to_datetime(ratings['year'])
# ratings['year'] = ratings['year'].dt.year
# ratings['date'] = pd.to_datetime(joined['date'])
# ratings['year'] = ratings['date'].dt.year
# pivot = ratings.pivot_table(values='rating', index='year', columns='genres', aggfunc='mean')
# c = pivot.loc[1996:2019]['Comedy'].sort_values()
# print(c)

orders = pd.read_csv('data/orders.csv', sep=';', header=1, names=['data','order ID', 'ID buyer', 'status', 'pay', 'cancel', 'otgruzhen', 'Product_ID', 'count'])
products = pd.read_csv('data/products.csv', sep=';')
# print(orders)
# Product_ID;Name;Price;CURRENCY    Дата создания;Order ID;ID Покупателя;Статус;Оплачен;Отменен;Отгружен;ID товара;Количество

orders_products = orders.merge(products, on='Product_ID',how='left')
# print(orders_products[orders_products['cancel'] == 'Да']['Name'])

orders_products['Profit'] = orders_products['Price'] * orders_products['count']
print(orders_products[orders_products['pay'] == 'Да'].groupby('ID buyer')['Profit'].sum().sort_values(ascending=False))