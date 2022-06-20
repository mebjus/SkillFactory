"""Игра угадай число
Компьютер сам загадывает и сам угадывает число
"""
import numpy as np


def random_predict(number, N=100) -> int:
    """Рандомно угадываем число
    Args:
        number (int, optional): Загаданное число. Defaults to 1.
    Return
        int: Число попыток
    """
    count = 0
    a, b = 1, N + 1
    while True:
        count += 1
        predict_number = int(np.mean([a, b]))
        if number > predict_number:
            a = predict_number
        elif number < predict_number:
            b = predict_number
        elif number == predict_number:
            break  # выход из цикла если угадали
    return count


def score_game(N=100, M=1000) -> int:
    """За какое количство попыток в среднем за 1000 подходов угадывает наш алгоритм
    Args:
        random_predict ([type]): функция угадывания
    Returns:
        int: среднее количество попыток
    """
    count_ls = []
    # np.random.seed(1)  # фиксируем сид для воспроизводимости
    random_array = np.random.randint(1, N + 1, size=(M))  # загадали список чисел
    for number in random_array:
        count_ls.append(random_predict(number, N))
    score = int(np.mean(count_ls))
    print(f"Ваш алгоритм угадывает число в среднем за:{score} попыток")
    return score


if __name__ == "__main__":
    # RUN
    # запускаем с параметрами: первый N - диапазон рандомных чисел
    # второй параметр M - количество попыток угадывания
    # по умолчанию: N = 100, M = 1000

    score_game()
