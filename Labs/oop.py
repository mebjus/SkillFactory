# class DepartmentReport():
#     def __init__(self, company_name):
#         self.deals = []
#         self.company_name = company_name
#
#     def add_revenue(self, revenues):
#         if not hasattr(self, 'revenues'):
#             self.revenues  = []
#         self.revenues.append(revenues)
#
#     def average_revenue(self):
#         return (f'Average department revenue for {self.company_name}: {round(sum(self.revenues)/len(self.revenues))}')
#
# report = DepartmentReport("Danon")
# report.add_revenue(1_000_000)
# report.add_revenue(400_000)
#
# print(report.average_revenue())


# class SalesReport():
#     def __init__(self, employee_name):
#         self.deals = []
#         self.employee_name = employee_name
#
#     def add_deal(self, company, amount):
#         self.deals.append({'company': company, 'amount': amount})
#
#     def total_amount(self):
#         return sum([deal['amount'] for deal in self.deals])
#
#     def average_deal(self):
#         return self.total_amount()/len(self.deals)
#
#     def all_companies(self):
#         return list(set([deal['company'] for deal in self.deals]))
#
#     def print_report(self):
#         print("Employee: ", self.employee_name)
#         print("Total sales:", self.total_amount())
#         print("Average sales:", self.average_deal())
#         print("Companies:", self.all_companies())
#
#
# report = SalesReport("Ivan Semenov")
# report2 = SalesReport("oleg iii")
#
# report2.add_deal("PepsiCo", 120_000)
# report.add_deal("SkyEng", 250_000)
# report.add_deal("PepsiCo", 20_000)
#
# report.print_report()
# # => Employee:  Ivan


# class User():
#     # Базовые данные
#     def __init__(self, email, password, balance):
#         self.email = email
#         self.password = password
#         self.balance = balance
#
#     def login(self, email, password):
#         return (self.email == email) and (self.password == password)
#
#     def update_balance(self, amount):
#         self.balance +=  amount
#
# user = User("gosha@roskino.org", "qwerty", 20_000)
# user.login("gosha@roskino.org", "qwerty123")
# # => False
# user.login("gosha@roskino.org", "qwerty")
# # => True
# user.update_balance(200)
# user.update_balance(-500)
# print(user.balance)
# # => 19700

import statistics

import pickle
from datetime import datetime
from os import path

class Dumper():
    def __init__(self, archive_dir="archive/"):
        self.archive_dir = archive_dir

    def dump(self, data):
        # Библиотека pickle позволяет доставать и класть объекты в файл
        with open(self.get_file_name(), 'wb') as file:
            pickle.dump(data, file)

    def load_for_day(self, day):
        file_name = path.join(self.archive_dir, day + ".pkl")
        with open(file_name, 'rb') as file:
            sets = pickle.load(file)
        return sets

    # возвращает корректное имя для файла
    def get_file_name(self):
        today = datetime.now().strftime("%y-%m-%d")
        return path.join(self.archive_dir, today + ".pkl")



class OwnLogger():
    def __init__(self):
        self.buffer = {} #буфер, в котором хранятся сообщения

    def log(self, message, level):
        self.buffer[level] = message
        self.buffer['all'] = message

    def show_last(self, level='all'):
        if level in self.buffer:
            return self.buffer[level]
        else:
            return None


data = {
    'perfomance': [10, 20, 10],
    'clients': {"Romashka": 10, "Vector": 34}
}


dumper = Dumper()
dumper.dump(data)

# Восстановим для сегодняшней даты
file_name = datetime.now().strftime("%y-%m-%d")
restored_data = dumper.load_for_day(file_name)
##########
logger = OwnLogger()
logger.log("System started", "info")
logger.show_last("error")
# => None
# Некоторые интерпретаторы Python могут не выводить None, тогда в этой проверке у вас будет пустая строка
logger.log("Connection instable", "warning")
logger.log("Connection lost", "error")

logger.show_last()
# => Connection lost
logger.show_last("info")
# => System started

import os
start_path = os.getcwd()
print(start_path)
print(os.listdir())
