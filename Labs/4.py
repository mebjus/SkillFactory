import pandas as pd

def get_experience(arg):
    """
    Напишите функцию get_experience(arg), аргументом которой является строка столбца с опытом работы.
    Функция должна возвращать опыт работы в месяцах. Не забудьте привести результат к целому числу.
    """
    a = 0
    exclude_list1 = ['месяца', 'месяцев', 'месяц']
    exclude_list2 = ['лет', 'года', 'год']
    arg = arg.split()
    arg_end = arg[-1]
    if (arg_end in exclude_list1) and (arg[-3] in exclude_list2):
        a = int(arg[-4])*12 +int(arg[-2])
    if arg_end in exclude_list2:
        a = int(arg[-2])*12
    if (arg_end in exclude_list1) and (arg[-3] not in exclude_list2):
        a = int(arg[-2])
    print(a)

    return a

if __name__ == '__main__':
    experience_col = pd.Series([
        'Опыт работы 8 лет 3 месяца',
        'Опыт работы 3 года 5 месяцев',
        'Опыт работы 1 год 9 месяцев',
        'Опыт работы 3 месяца',
        'Опыт работы 6 лет'
        ])
    experience_month = experience_col.apply(get_experience)
