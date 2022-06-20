import numpy as np
import pandas as pd

# a=np.int8(25)
# # print(np.iinfo(np.int64))
#
# # print(*sorted(map(str, set(np.sctypeDict.values()))), sep='\n')
#
# # a = np.uint8(-456)
# # print(a)
#
# arr, step = np.linspace(-6, 21, 60, endpoint = False, retstep=True)
# print(step)



# def get_chess(a):
#     arr = np.zeros(a*a, dtype=np.float16)
#     arr.shape=(a,a)
#     arr[::2, 1::2] = 1
#     arr[1::2, ::2] = 1
#     return arr

# def shuffle_seed(array):
#
#     seed_ = np.random.randint(0,2**32, dtype=np.int64)
#     np.random.seed(seed_)
#     new_array = np.random.permutation(array)
#     return (new_array, seed_)
#
# array = [1, 2, 3, 4, 5]
# shuffle_seed(array)

# Напишите функцию any_normal, которая принимает на вход неограниченное число векторов через запятую.
# Гарантируется, что все векторы, которые передаются, одинаковой длины.
#
# Функция возвращает True, если есть хотя бы одна пара перпендикулярных векторов. Иначе возвращает False.
#
# Пример:
#
# vec1 = np.array([2, 1])
# vec2 = np.array([-1, 2])
# vec3 = np.array([3,4])
# print(any_normal(vec1, vec2, vec3))
# # True

# def any_normal(*vectors):
#     d = list()
#     for i in range(len(vectors)):
#         for j in range(i):
#             d.append(np.dot(vectors[i], vectors[j]))
#     return (0 in d)
#
#
# vec1 = np.array([2, 1])
# vec2 = np.array([-1, 2])
# vec3 = np.array([3,4])
#
# print(any_normal(vec1, vec2, vec3))

#
# Напишите функцию get_loto(num), генерирующую трёхмерный массив случайных целых чисел от 1 до 100 (включительно).
# Это поля для игры в лото.
#
# Трёхмерный массив должен состоять из таблиц чисел формы 5х5, то есть итоговая форма — (num, 5, 5).
#
# Функция возвращает полученный массив.
#
# Пример:

# array = np.random.randint(1,101, size=(num,5,5))

# def get_unique_loto(num):
#     sample = np.arange(1, 101)
#     arr=np.array([], dtype=np.int16)
#     for i in range(num):
#         a = np.random.choice(sample, size=(5,5), replace=False)
#         arr = np.append(arr, a)
#     arr.shape=(num, 5, 5)
#     return print(arr)
#
# get_unique_loto(2)

# def get_unique_loto(num):
#     sample = np.arange(1, 101)
#     res = list()
#     for i in range(num):
#         res.append(np.random.choice(sample, replace=False, size=(5, 5)))
#     res = np.array(res)
#     return res


print(pd.__name__)
countries = pd.Series(
    data = ['Англия', 'Канада', 'США', 'Россия', 'Украина', 'Беларусь', 'Казахстан'],
    index = ['UK', 'CA', 'US', 'RU', 'UA', 'BY', 'KZ'],
    name = 'countries'
)
display(countries)