
f = open('data/test.txt', 'a', encoding='utf8') # открываем файл на дозапись

sequence = ["other string\n", "123\n", "test test\n"]
f.writelines(sequence) # берет строки из sequence и записывает в файл (без переносов)
f.close()

f = open('data/test.txt', 'r', encoding='utf8')
print(f.readlines()) # считывает все строки в список и возвращает список
f.close()

with open("data/test.txt", 'rb') as f:
    a = f.read(10)
    b = f.read(23)
