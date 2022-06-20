import pandas as pd
import numpy as np

student_data = pd.read_csv('data/students_performance.csv', sep=',')
# mask = student_dat['race/ethnicity'] == 'A'
# # mask2 = student_dat['race/ethnicity'] == 'C'
#
# print(student_dat[mask])
#
# # value_counts(normalize=True)

# a = student_data[student_data['race/ethnicity'] == 'group A']['writing score'].median()
# b = student_data[student_data['race/ethnicity'] == 'group C']['writing score'].mean()
# print(abs(round(a-b)))