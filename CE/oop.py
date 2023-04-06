from readdf import *

a = Readdf()
a.fo(own=False)
print(a[5]['Клиент'])

df = a.df
print(df)

####################
# with pd.ExcelWriter('data/tst.xlsx') as writer:
#     df.to_excel(writer, sheet_name="итоги")
