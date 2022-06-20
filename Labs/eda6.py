import  pandas as pd
import matplotlib.pyplot as plt
import logging
import os
from comet_ml import Experiment


df = pd.read_csv('data/KaggleV2-May-2016.csv')
encoded_columns = pd.get_dummies(df, columns=['Gender', 'Neighbourhood'])

# print(df['No-show'].value_counts(True))
# print(df[df['No-show'] == 'Yes']['Gender'].value_counts(True))

plt.figure(figsize=(8, 4))
df['No-show'].hist()
plt.title('Соотношение людей, которые пропускают приемы, к тем, кто этого не делает');
plt.show()
# logging.FileHandler('log_file.log')
# logging.basicConfig(format="%(levelname)s: %(asctime)s: %(message)s", level=logging.DEBUG)
# logging.info('Проверка')


# Функция для создания лог-файла и записи в него информации
def get_logger(path, file):
    """[Создает лог-файл для логирования в него]
    Аргументы:
        path {string} -- путь к директории
        file {string} -- имя файла
     Возвращает:
        [obj] -- [логер]
    """
    # проверяем, существует ли файл
    log_file = os.path.join(path, file)

    # если  файла нет, создаем его
    if not os.path.isfile(log_file):
        open(log_file, "w+").close()

    # поменяем формат логирования
    file_logging_format = "%(levelname)s: %(asctime)s: %(message)s"

    # конфигурируем лог-файл
    logging.basicConfig(level=logging.INFO,
                        format=file_logging_format)
    logger = logging.getLogger()

    # создадим хэнлдер для записи лога в файл
    handler = logging.FileHandler(log_file)

    # установим уровень логирования
    handler.setLevel(logging.INFO)

    # создадим формат логирования, используя file_logging_format
    formatter = logging.Formatter(file_logging_format)
    handler.setFormatter(formatter)

    # добавим хэндлер лог-файлу
    logger.addHandler(handler)
    return logger


logger = get_logger(path="logs/", file="data.logs")
logger.info("Data")

logger.info("Data shape {}".format(df.shape))
logger.info("Percentage of women: {}".format(df[df['No-show'] == 'Yes']['Gender'].value_counts(True)[0]))
logger.info("Percentage of men: {}".format(df[df['No-show'] == 'Yes']['Gender'].value_counts(True)[1]))

# Создайте эксперимент с помощью вашего API ключа
experiment = Experiment(
    api_key="vpUQ9NEGFGMpz4NRt9dS7QxVL",
    project_name="medical",
    workspace="mebjus",
)
